/* Certificate Reader — Excel task pane.
   Reads dropped certificates with Claude, matches them to a property on the Properties sheet and writes rows to the
   Certificates sheet of the open tracker. Uncertain results are shown for a quick check before saving. */
"use strict";

const MODEL = "claude-opus-5-5";
const API = "https://api.anthropic.com/v1/messages";
const FIRST = 5, UNIT_LAST = 504, CERT_LAST = 3004, TYPE_FIRST = 5, TYPE_LAST = 29;
const AUTO_SCORE = 60, AUTO_MARGIN = 15, PARALLEL = 3;
const PTYPES = ["Residential let", "HMO", "Block / common parts", "Commercial", "Mixed-use"];
const FLAGS = [["gas", "Gas"], ["lift", "Lift"], ["aircon", "Air con >12kW"], ["over11m", "Over 11m"], ["pre2000", "Built pre-2000"], ["licence", "Council licence"]];
const READABLE = /\.(pdf|jpe?g|png|webp|gif|heic|heif)$/i;
const KEY = "gc-cert-reader-key";

const $ = (id) => document.getElementById(id);
let busy = 0, queue = [], stats = { saved: 0, review: 0, skipped: 0, error: 0 };

/* ------------------------------------------------------------ workbook */
async function loadLists() {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const u = wb.worksheets.getItem("Properties").getRange(`B${FIRST}:N${UNIT_LAST}`).load("values");
    const t = wb.worksheets.getItem("Certificate Types").getRange(`B${TYPE_FIRST}:B${TYPE_LAST}`).load("values");
    const c = wb.worksheets.getItem("Certificates").getRange(`B${FIRST}:D${CERT_LAST}`).load("values");
    await ctx.sync();
    const units = u.values.map((r) => ({ building: String(r[0] || "").trim(), unit: String(r[1] || "").trim(), postcode: String(r[2] || ""), active: r[12] !== "No" }))
      .filter((x) => x.building).map((x) => ({ ...x, name: displayName(x.building, x.unit) }));
    const types = t.values.map((r) => String(r[0] || "").trim()).filter(Boolean);
    const existing = new Set(c.values.filter((r) => r[0] && r[1] && r[2]).map((r) => `${r[0]}|${r[1]}|${Math.floor(Number(r[2]))}`));
    return { units, types, existing };
  });
}

async function writeCert(row) {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem("Certificates");
    const r = ws.getRange(`B${FIRST}:D${CERT_LAST}`).load("values");
    await ctx.sync();
    const i = r.values.findIndex((v) => v.every((x) => x === "" || x === null));
    if (i < 0) throw new Error("The Certificates sheet is full (3,000 rows).");
    const n = FIRST + i;
    ws.getRange(`B${n}:M${n}`).values = [[row.property, row.type, serial(row.issue), serial(row.expiry), row.result || "", "", "",
      row.company || "", row.number || "", row.file || "", row.notes || "", "Reader"]];
    await ctx.sync();
    return n;
  });
}

async function addProperty(p) {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem("Properties");
    const r = ws.getRange(`B${FIRST}:B${UNIT_LAST}`).load("values");
    await ctx.sync();
    const i = r.values.findIndex((v) => !v[0]);
    if (i < 0) throw new Error("The Properties sheet is full (500 rows).");
    const n = FIRST + i;
    ws.getRange(`B${n}:K${n}`).values = [[p.building, p.unit || "", p.postcode || "", p.type,
      ...FLAGS.map(([k]) => (p.flags[k] ? "Yes" : "No"))]];
    ws.getRange(`N${n}`).values = [["Yes"]];
    await ctx.sync();
    return displayName(p.building, p.unit);
  });
}

/* ------------------------------------------------------------ matching (same rules as the desktop reader) */
const displayName = (b, u) => (u ? `${b} – ${u}` : b);
const PC = /\b([A-Z]{1,2}\d[A-Z\d]?)\s*(\d[A-Z]{2})\b/i;
function normPostcode(s) { const m = PC.exec(String(s || "").toUpperCase()); return m ? m[1] + m[2] : ""; }
function norm(s) {
  return String(s || "").toLowerCase().replace(/\broad\b/g, "rd").replace(/\bstreet\b/g, "st").replace(/\bavenue\b/g, "ave")
    .replace(/\bcourt\b/g, "ct").replace(/\bapartment\b|\bapt\b/g, "flat").replace(/[^a-z0-9 ]/g, " ").split(/\s+/).filter(Boolean).join(" ");
}
const numbers = (s) => new Set(norm(s).match(/\b\d+[a-z]?\b/g) || []);
function unitKey(s) {
  s = norm(s);
  if (!s || s === "whole" || s === "whole building" || s === "house") return "whole";
  if (/common|communal|landlords? supply|block/.test(s)) return "common";
  if (/shop|retail|commercial/.test(s)) return "shop";
  const m = /(?:flat|unit|room)\s*([0-9]+[a-z]?)/.exec(s) || /\b([0-9]+[a-z]?)\b/.exec(s);
  return m ? `no:${m[1]}` : s;
}
function match(units, address, unit, postcode) {
  const pc = normPostcode(postcode) || normPostcode(address), nums = numbers(address);
  const words = new Set(norm(address).split(" ").filter((w) => !nums.has(w))), uk = unitKey(unit);
  const scored = units.filter((u) => u.active).map((u) => {
    let s = 0;
    const upc = normPostcode(u.postcode);
    if (pc && upc) s += pc === upc ? 40 : pc.slice(0, -3) === upc.slice(0, -3) ? 15 : -30;
    const bn = numbers(u.building), base = (x) => x.replace(/[a-z]$/, "");
    const shared = [...bn].some((x) => nums.has(x) || [...nums].some((y) => base(y) === base(x)));
    if (shared) s += 25; else if (bn.size && nums.size) s -= 15;
    const bw = norm(u.building).split(" ").filter((w) => w && !bn.has(w));
    if (bw.length) s += (20 * bw.filter((w) => words.has(w)).length) / bw.length;
    const pk = unitKey(u.unit);
    if (pk === uk) s += 25;
    else if (uk === "whole" && pk.startsWith("no:") && nums.has(pk.slice(3))) s += 25; // "610a Holloway Road" -> Flat 610a
    else if (uk === "whole" && pk === "common") s += 5; else if (pk !== "whole" && uk !== "whole") s -= 20;
    return [s, u];
  }).sort((a, b) => b[0] - a[0]);
  if (!scored.length) return { unit: null, score: 0, second: 0 };
  return { unit: scored[0][1], score: scored[0][0], second: scored[1] ? scored[1][0] : 0 };
}

/* ------------------------------------------------------------ Claude */
const PROMPT = (types) => `You are reading documents from a UK letting and property management office. Find every property
compliance certificate in the attached document (one PDF can hold several, e.g. a fire alarm and an emergency
lighting certificate). For each one, fill in the fields exactly as the document states them.

Certificate type — choose the closest from this list, or "Other" if none fits:
${types.map((t) => "- " + t).join("\n")}

How to classify common documents:
- Landlord Gas Safety Record / CP12 / Domestic Gas Safety Record -> Gas Safety Record (CP12)
- Electrical Installation Condition Report (EICR), and Electrical Installation Certificate (EIC) or NICEIC
  installation certificate for a flat or house -> EICR (electrical). For an EIC, the issue date is the completion
  date and the expiry is the "next inspection" date written on it.
- EICR for a shop, office or other non-domestic premises -> Fixed wire test (commercial EICR)
- Energy Performance Certificate -> EPC (result = rating band and score, e.g. "C (72)"; expiry = "valid until")
- Fire risk assessment -> Fire risk assessment (expiry = the stated review date, if any)
- Fire alarm / fire detection inspection or servicing certificate (BS 5839) -> Fire alarm service
- Emergency lighting certificate (BS 5266) -> Emergency lighting test
- Portable appliance testing -> PAT testing
- HMO licence -> HMO licence; selective or additional licence -> Selective / additional licence
  (issue = licence start, expiry = licence end)
- Council application forms, payment pages, floor plans, invoices, letters and photos that are not a
  certificate are NOT certificates: return an empty list and set is_certificate to false.

Rules:
- Dates as YYYY-MM-DD. If a date is not on the document, use "". Never invent or calculate a date — the tracker
  works out due dates itself when the document doesn't state one.
- Address: the property inspected (not the landlord's, agent's or contractor's address). Put the flat, unit, shop
  or "Common parts" / "Landlord supply" in unit, and the building number and street in address.
- result: Satisfactory, Unsatisfactory, Pass, Fail, Remedial work needed or Issued (for licences and assessments),
  or the EPC band and score.
- confidence: high only when type, address and issue date are all clearly printed; otherwise medium or low.`;

function schema(types) {
  const str = { type: "string" };
  const fields = ["type_as_written", "address", "unit", "postcode", "issue_date", "expiry_or_next_due", "result", "remedial_work",
    "company", "engineer", "certificate_number", "pages"];
  const props = { type: { type: "string", enum: [...types, "Other"] }, confidence: { type: "string", enum: ["high", "medium", "low"] } };
  fields.forEach((f) => (props[f] = str));
  return {
    type: "object", additionalProperties: false, required: ["is_certificate", "certificates"],
    properties: {
      is_certificate: { type: "boolean" },
      certificates: { type: "array", items: { type: "object", additionalProperties: false, required: ["type", "confidence", ...fields], properties: props } },
    },
  };
}

async function readWithClaude(file, types) {
  let blob = file, name = file.name;
  if (/\.hei[cf]$/i.test(name)) {
    if (!window.heic2any) throw new Error("HEIC converter didn't load — check the internet connection.");
    blob = await window.heic2any({ blob: file, toType: "image/jpeg", quality: 0.9 });
    name = name.replace(/\.hei[cf]$/i, ".jpg");
  }
  if (blob.size > 30 * 1024 * 1024) throw new Error("File is over 30 MB — split it into smaller PDFs.");
  const data = await toBase64(blob);
  const isPdf = /\.pdf$/i.test(name);
  const media = isPdf ? "application/pdf" : blob.type || "image/jpeg";
  const block = isPdf ? { type: "document", source: { type: "base64", media_type: media, data } }
    : { type: "image", source: { type: "base64", media_type: media, data } };
  const res = await fetch(API, {
    method: "POST",
    headers: {
      "content-type": "application/json", "x-api-key": localStorage.getItem(KEY) || "", "anthropic-version": "2026-10-21",
      "anthropic-beta": "server-side-fallback-2026-07-01", "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: MODEL, max_tokens: 16000, fallbacks: "default",
      output_config: { effort: "medium", format: { type: "json_schema", schema: schema(types) } },
      messages: [{ role: "user", content: [block, { type: "text", text: PROMPT(types) + `\nFile name (may help, may be meaningless): ${file.name}` }] }],
    }),
  });
  const body = await res.json().catch(() => ({}));
  if (res.status === 401) throw new Error("The API key isn't valid — change it in Settings.");
  if (res.status === 429) throw new Error("Too many at once — wait a minute and drop it again.");
  if (!res.ok) throw new Error((body.error && body.error.message) || `Claude API error ${res.status}`);
  if (body.stop_reason === "refusal") throw new Error("The document could not be read.");
  if (body.stop_reason === "max_tokens") throw new Error("The document is too long to read in one go.");
  const text = (body.content || []).find((b) => b.type === "text");
  return JSON.parse(text ? text.text : "{}");
}

function toBase64(blob) {
  return new Promise((ok, fail) => {
    const r = new FileReader();
    r.onload = () => ok(String(r.result).split(",")[1]);
    r.onerror = () => fail(r.error);
    r.readAsDataURL(blob);
  });
}

/* ------------------------------------------------------------ dates */
function serial(iso) {
  if (!iso || !/^\d{4}-\d{2}-\d{2}/.test(iso)) return "";
  const [y, m, d] = iso.slice(0, 10).split("-").map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}
const ukDate = (iso) => (iso ? iso.split("-").reverse().join("/") : "—");

/* ------------------------------------------------------------ processing */
function addFiles(files) {
  for (const f of files) queue.push(f);
  pump();
}

function pump() {
  while (busy < PARALLEL && queue.length) {
    const f = queue.shift();
    busy++;
    handle(f).finally(() => { busy--; pump(); });
  }
  showSummary();
}

async function handle(file) {
  const card = newCard(file.name);
  if (!READABLE.test(file.name)) return finish(card, "skipped", "Not a PDF or photo — skipped.");
  try {
    const { units, types, existing } = await loadLists();
    const data = await readWithClaude(file, types);
    const items = (data.certificates || []).filter((c) => c.type !== "Other" || c.issue_date);
    if (!data.is_certificate || !items.length) return finish(card, "skipped", "Not a certificate (form, plan, letter or photo) — skipped.");
    card.remove();
    for (const c of items) {
      const sub = newCard(`${file.name} · ${c.type_as_written || c.type}`);
      const m = match(units, c.address, c.unit, c.postcode);
      const reasons = [];
      if (!types.includes(c.type)) reasons.push("certificate type");
      if (!c.issue_date) reasons.push("issue date");
      if (!m.unit || m.score < AUTO_SCORE || m.score - m.second < AUTO_MARGIN) reasons.push("which property");
      if (c.confidence === "low") reasons.push("hard to read");
      const item = { ...c, file: file.name, suggested: m.unit && m.score > 20 ? m.unit.name : "" };
      if (!reasons.length && existing.has(`${m.unit.name}|${c.type}|${serial(c.issue_date)}`)) {
        finish(sub, "skipped", `Already in the tracker: ${m.unit.name} · ${c.type} · ${ukDate(c.issue_date)}`);
      } else if (reasons.length) {
        reviewForm(sub, item, units, types, "Please check: " + reasons.join(", ") + ".");
      } else {
        await save(item, m.unit.name);
        finish(sub, "saved", `Saved: ${m.unit.name} · ${c.type} · ${ukDate(c.issue_date)}${c.expiry_or_next_due ? " → " + ukDate(c.expiry_or_next_due) : ""}`);
      }
    }
  } catch (e) {
    finish(card, "error", e.message || String(e));
  }
}

async function save(c, propertyName) {
  const notes = [c.engineer && `Engineer: ${c.engineer}`, c.remedial_work && `Remedial: ${c.remedial_work}`].filter(Boolean).join("; ");
  await writeCert({ property: propertyName, type: c.type, issue: c.issue_date, expiry: c.expiry_or_next_due, result: c.result,
    company: c.company, number: c.certificate_number, file: c.file, notes });
}

/* ------------------------------------------------------------ UI */
function newCard(name) {
  const d = document.createElement("div");
  d.className = "card working";
  d.innerHTML = `<div class="name"></div><div class="msg"><span class="spin"></span>Reading…</div>`;
  d.querySelector(".name").textContent = name;
  $("list").prepend(d);
  return d;
}

function finish(card, status, msg) {
  card.className = `card ${status}`;
  card.querySelector(".msg").textContent = msg;
  stats[status] = (stats[status] || 0) + 1;
  showSummary();
}

function showSummary() {
  const waiting = busy + queue.length;
  $("summary").textContent = (waiting ? `Reading ${waiting}… ` : "") +
    `Saved ${stats.saved} · To check ${stats.review} · Skipped ${stats.skipped} · Errors ${stats.error}`;
}

function reviewForm(card, c, units, types, why) {
  card.className = "card review";
  stats.review++;
  showSummary();
  const opt = (v, sel) => `<option${v === sel ? " selected" : ""}>${esc(v)}</option>`;
  const props = units.filter((u) => u.active).map((u) => u.name).sort();
  card.querySelector(".msg").innerHTML = `
    <div>${esc(why)}</div>
    <div class="msg">Read as: ${esc([c.address, c.unit, c.postcode].filter(Boolean).join(", ") || "no address found")}</div>
    <label>Property</label>
    <select class="f-prop"><option value="">— choose —</option>${props.map((p) => opt(p, c.suggested)).join("")}<option value="__new">＋ New property…</option></select>
    <div class="newprop" hidden>
      <div class="row"><div><label>Building / address</label><input class="n-b" value="${esc(c.address)}"></div><div><label>Unit</label><input class="n-u" value="${esc(c.unit)}"></div></div>
      <div class="row"><div><label>Postcode</label><input class="n-p" value="${esc(c.postcode)}"></div>
        <div><label>Type</label><select class="n-t">${PTYPES.map((t) => opt(t, "")).join("")}</select></div></div>
      <div class="flags">${FLAGS.map(([k, l]) => `<label><input type="checkbox" data-k="${k}">${l}</label>`).join("")}</div>
    </div>
    <label>Certificate</label><select class="f-type">${types.map((t) => opt(t, c.type)).join("")}</select>
    <div class="row"><div><label>Issue date</label><input class="f-issue" type="date" value="${esc(c.issue_date)}"></div>
      <div><label>Expiry / next due (if stated)</label><input class="f-exp" type="date" value="${esc(c.expiry_or_next_due)}"></div></div>
    <div class="row"><div><label>Result</label><input class="f-res" value="${esc(c.result)}"></div>
      <div><label>Certificate no.</label><input class="f-no" value="${esc(c.certificate_number)}"></div></div>
    <label>Company</label><input class="f-co" value="${esc(c.company)}">
    <button class="f-save">Save to tracker</button> <button class="light f-skip">Skip</button>`;
  const q = (s) => card.querySelector(s);
  q(".f-prop").onchange = () => (q(".newprop").hidden = q(".f-prop").value !== "__new");
  q(".f-skip").onclick = () => { stats.review--; finish(card, "skipped", "Skipped."); };
  q(".f-save").onclick = async () => {
    const btn = q(".f-save");
    btn.disabled = true;
    try {
      let prop = q(".f-prop").value;
      if (!prop) throw new Error("Choose the property.");
      if (!q(".f-issue").value) throw new Error("Add the issue date.");
      if (prop === "__new") {
        const flags = {};
        card.querySelectorAll(".flags input").forEach((i) => (flags[i.dataset.k] = i.checked));
        if (!q(".n-b").value.trim()) throw new Error("Add the building / address.");
        prop = await addProperty({ building: q(".n-b").value.trim(), unit: q(".n-u").value.trim(), postcode: q(".n-p").value.trim(), type: q(".n-t").value, flags });
      }
      const item = { ...c, type: q(".f-type").value, issue_date: q(".f-issue").value, expiry_or_next_due: q(".f-exp").value,
        result: q(".f-res").value, certificate_number: q(".f-no").value, company: q(".f-co").value };
      await save(item, prop);
      stats.review--;
      finish(card, "saved", `Saved: ${prop} · ${item.type} · ${ukDate(item.issue_date)}`);
    } catch (e) {
      btn.disabled = false;
      alert(e.message || String(e));
    }
  };
}

const esc = (s) => String(s || "").replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));

/* folders dropped onto the pane: walk them */
async function entriesToFiles(items) {
  const out = [];
  const walk = (entry) => new Promise((ok) => {
    if (entry.isFile) return entry.file((f) => { out.push(f); ok(); }, () => ok());
    const reader = entry.createReader(), all = [];
    const next = () => reader.readEntries(async (batch) => {
      if (!batch.length) { for (const e of all) await walk(e); return ok(); }
      all.push(...batch); next();
    }, () => ok());
    next();
  });
  for (const it of items) {
    const e = it.webkitGetAsEntry && it.webkitGetAsEntry();
    if (e) await walk(e); else if (it.getAsFile) out.push(it.getAsFile());
  }
  return out.filter((f) => f && !f.name.startsWith("."));
}

function needKey() { $("setup").hidden = !!localStorage.getItem(KEY); }

async function showOrg() {  // company name comes from the open workbook (Setup!C5), never from the published code
  try {
    await Excel.run(async (ctx) => {
      const r = ctx.workbook.worksheets.getItem("Setup").getRange("C5").load("values");
      await ctx.sync();
      $("org").textContent = String(r.values[0][0] || "");
    });
  } catch (e) { /* not a tracker workbook: leave it blank */ }
}

Office.onReady(() => {
  needKey();
  showOrg();
  $("saveKey").onclick = () => { const k = $("key").value.trim(); if (k) localStorage.setItem(KEY, k); $("key").value = ""; needKey(); };
  $("changeKey").onclick = () => { localStorage.removeItem(KEY); needKey(); };
  const drop = $("drop");
  drop.onclick = () => $("pick").click();
  $("pick").onchange = (e) => { addFiles([...e.target.files]); e.target.value = ""; };
  drop.ondragover = (e) => { e.preventDefault(); drop.classList.add("over"); };
  drop.ondragleave = () => drop.classList.remove("over");
  drop.ondrop = async (e) => {
    e.preventDefault();
    drop.classList.remove("over");
    if (!localStorage.getItem(KEY)) return needKey();
    addFiles(await entriesToFiles([...e.dataTransfer.items]));
  };
  showSummary();
});
