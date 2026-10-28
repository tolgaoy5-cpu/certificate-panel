/* Certificate Reader — Excel task pane.
   Reads dropped certificates with the OpenAI API, matches them to a property on the Properties sheet and writes rows to the
   Certificates sheet of the open tracker. Uncertain results are shown for a quick check before saving. */
"use strict";

const MODEL = "gpt-6.1-sol";
const API = "https://api.openai.com/v1/responses";
const FIRST = 5, UNIT_LAST = 504, CERT_LAST = 3004, TYPE_FIRST = 5, TYPE_LAST = 29;
const AUTO_SCORE = 60, AUTO_MARGIN = 15, PARALLEL = 3;
const PTYPES = ["Residential let", "HMO", "Block / common parts", "Commercial", "Mixed-use"];
const FLAGS = [["gas", "Gas"], ["lift", "Lift"], ["aircon", "Air con >12kW"], ["over11m", "Over 11m"], ["pre2000", "Built pre-2000"], ["licence", "Council licence"]];
const READABLE = /\.(pdf|jpe?g|png|webp|gif|heic|heif)$/i;
const KEY = "gc-cert-reader-key";
const ARCHIVE = "https://127.0.0.1:8765";  // Certificate Inbox service on this Mac: keeps a copy of every certificate

const $ = (id) => document.getElementById(id);
let busy = 0, queue = [], stats = { saved: 0, review: 0, skipped: 0, error: 0 }, archiveOn = false;

/* ------------------------------------------------------------ workbook */
async function loadLists() {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const u = wb.worksheets.getItem("Properties").getRange(`B${FIRST}:N${UNIT_LAST}`).load("values");
    const t = wb.worksheets.getItem("Certificate Types").getRange(`B${TYPE_FIRST}:B${TYPE_LAST}`).load("values");
    const c = wb.worksheets.getItem("Certificates").getRange(`B${FIRST}:D${CERT_LAST}`).load("values");
    await ctx.sync();
    const units = u.values.map((r) => ({ building: String(r[0] || "").trim(), unit: String(r[1] || "").trim(), postcode: String(r[2] || ""),
      type: String(r[3] || ""), active: r[12] !== "No" }))
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
    if (r.values.some((v) => v[0] === row.property && v[1] === row.type && Math.floor(Number(v[2])) === serial(row.issue))) return 0; // already there
    const i = r.values.findIndex((v) => v.every((x) => x === "" || x === null));
    if (i < 0) throw new Error("The Certificates sheet is full (3,000 rows).");
    const n = FIRST + i;
    ws.getRange(`B${n}:M${n}`).values = [[row.property, row.type, serial(row.issue), serial(row.expiry), row.result || "", "", "",
      row.company || "", row.number || "", row.file || "", row.notes || "", "Reader"]];
    if (row.link) ws.getRange(`K${n}`).hyperlink = { address: row.link, textToDisplay: row.file };
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
  return String(s || "").toLowerCase().replace(/['’]/g, "").replace(/\broad\b/g, "rd").replace(/\bstreet\b/g, "st").replace(/\bavenue\b/g, "ave")
    .replace(/\bcourt\b/g, "ct").replace(/\bapartment\b|\bapt\b/g, "flat").replace(/[^a-z0-9 ]/g, " ").split(/\s+/).filter(Boolean).join(" ");
}
const numbers = (s) => new Set(norm(s).match(/\b\d+[a-z]?\b/g) || []);
const BUILDING_TYPES = new Set(["Fire risk assessment", "Fire alarm service", "Emergency lighting test", "Fire extinguisher service",
  "Fire door checks (communal)", "Lift thorough examination (LOLER)", "Asbestos survey / re-inspection", "Dry riser test"]);
function unitKey(s) {
  s = norm(s);
  if ((s.match(/\b\d+[a-z]?\b/g) || []).length > 1) return "multi"; // "Rooms 2, 3, 4 and 5": never one flat
  if (["flats", "all flats", "rooms", "all rooms", "building", "whole block"].includes(s)) return "common";
  if (!s || s === "whole" || s === "whole building" || s === "house") return "whole";
  if (/common|communal|landlords? supply|block/.test(s)) return "common";
  if (/shop|retail|commercial/.test(s)) return "shop";
  const m = /(?:flat|unit|room)\s*([0-9]+[a-z]?)/.exec(s) || /\b([0-9]+[a-z]?)\b/.exec(s);
  return m ? `no:${m[1]}` : s;
}
function match(units, address, unit, postcode, ctype) {
  const pc = normPostcode(postcode) || normPostcode(address), nums = numbers(address);
  const words = new Set(norm(address).split(" ").filter((w) => !nums.has(w)));
  const uk0 = unitKey(unit);
  const active = units.filter((u) => u.active), per = {}, hasCommon = new Set();
  active.forEach((u) => {
    per[norm(u.building)] = (per[norm(u.building)] || 0) + 1;
    if (unitKey(u.unit) === "common") hasCommon.add(norm(u.building));
  });
  const scored = active.map((u) => {
    let s = 0;
    const onlyUnit = per[norm(u.building)] === 1; // one unit in the building: its name can be written many ways
    // building-wide certificates belong to the common parts, where the building has them
    const uk = BUILDING_TYPES.has(ctype) && (uk0 === "whole" || uk0 === "multi") && hasCommon.has(norm(u.building)) ? "common" : uk0;
    const upc = normPostcode(u.postcode);
    if (pc && upc) s += pc === upc ? 40 : pc.slice(0, -3) === upc.slice(0, -3) ? 15 : -30;
    const bn = numbers(u.building), base = (x) => x.replace(/[a-z]$/, "");
    const shared = [...bn].some((x) => nums.has(x) || [...nums].some((y) => base(y) === base(x)));
    if (shared) s += 25; else if (bn.size && nums.size) s -= 15;
    const bw = norm(u.building).split(" ").filter((w) => w && !bn.has(w));
    if (bw.length) s += (20 * bw.filter((w) => words.has(w)).length) / bw.length;
    const pk = unitKey(u.unit);
    if (uk === "multi") s -= 20;
    else if (pk === uk) s += 25;
    else if (uk === "whole" && pk.startsWith("no:") && nums.has(pk.slice(3))) s += 25; // "610a Holloway Road" -> Flat 610a
    else if (uk === "whole" && pk === "common") s += 5;
    else if (onlyUnit) s += 10;
    else if (pk !== "whole" && uk !== "whole") s -= 20;
    if (ctype === "HMO licence" && u.type === "HMO") s += 15;
    return [s, u];
  }).sort((a, b) => b[0] - a[0]);
  if (!scored.length) return { unit: null, score: 0, second: 0 };
  return { unit: scored[0][1], score: scored[0][0], second: scored[1] ? scored[1][0] : 0 };
}

/* ------------------------------------------------------------ OpenAI */
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

async function readCertificate(file, types) {
  let blob = file, name = file.name;
  if (/\.hei[cf]$/i.test(name)) {
    if (!window.heic2any) throw new Error("HEIC converter didn't load — check the internet connection.");
    blob = await window.heic2any({ blob: file, toType: "image/jpeg", quality: 0.9 });
    name = name.replace(/\.hei[cf]$/i, ".jpg");
  }
  if (blob.size > 45 * 1024 * 1024) throw new Error("File is over 45 MB — split it into smaller PDFs.");
  const isPdf = /\.pdf$/i.test(name);
  const part = isPdf
    ? { type: "input_file", filename: name, file_data: "data:application/pdf;base64," + (await toBase64(blob)), detail: "high" }
    : { type: "input_image", image_url: await imageDataUrl(blob), detail: "high" };
  const res = await fetch(API, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: "Bearer " + (localStorage.getItem(KEY) || "") },
    body: JSON.stringify({
      model: MODEL, max_output_tokens: 16000,
      input: [{ role: "user", content: [part, { type: "input_text", text: PROMPT(types) + `\nFile name (may help, may be meaningless): ${file.name}` }] }],
      text: { format: { type: "json_schema", name: "certificates", schema: schema(types), strict: true } },
    }),
  });
  const body = await res.json().catch(() => ({}));
  if (res.status === 401) throw new Error("The API key isn't valid — change it in Settings.");
  if (res.status === 429) throw new Error("Too many at once, or the account is out of credit — wait a minute and try again.");
  if (!res.ok) throw new Error((body.error && body.error.message) || `OpenAI API error ${res.status}`);
  if (body.status === "incomplete") throw new Error("The document could not be read in full.");
  const parts = (body.output || []).flatMap((o) => o.content || []);
  if (parts.some((p) => p.type === "refusal")) throw new Error("The document could not be read.");
  const text = parts.find((p) => p.type === "output_text");
  return JSON.parse(text ? text.text : "{}");
}

/* photos: shrink to at most 2400px so requests stay small */
function imageDataUrl(blob) {
  return new Promise((ok, fail) => {
    const img = new Image(), url = URL.createObjectURL(blob);
    img.onload = () => {
      const k = Math.min(1, 2400 / Math.max(img.width, img.height));
      const c = document.createElement("canvas");
      c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      URL.revokeObjectURL(url);
      ok(c.toDataURL("image/jpeg", 0.88));
    };
    img.onerror = () => fail(new Error("This photo couldn't be opened."));
    img.src = url;
  });
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
/* An EPC is valid for exactly 10 years, so a missing issue date follows from "valid until". */
function epcIssue(c) {
  if (c.type !== "EPC" || c.issue_date || !/^\d{4}-\d{2}-\d{2}/.test(c.expiry_or_next_due || "")) return "";
  const [y, m, d] = c.expiry_or_next_due.slice(0, 10).split("-").map(Number);
  const next = new Date(Date.UTC(y, m - 1, d + 1));
  let issue = new Date(Date.UTC(next.getUTCFullYear() - 10, next.getUTCMonth(), next.getUTCDate()));
  if (issue.getUTCMonth() !== next.getUTCMonth()) issue = new Date(Date.UTC(next.getUTCFullYear() - 10, next.getUTCMonth(), next.getUTCDate() - 1)); // 29 Feb
  c.issue_date = issue.toISOString().slice(0, 10);
  return "Issue date worked out from the expiry (EPCs last 10 years)";
}
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
    const data = await readCertificate(file, types);
    const items = (data.certificates || []).filter((c) => c.type !== "Other" || c.issue_date);
    if (!data.is_certificate || !items.length) return finish(card, "skipped", "Not a certificate (form, plan, letter or photo) — skipped.");
    card.remove();
    for (const c of items) {
      const sub = newCard(`${file.name} · ${c.type_as_written || c.type}`);
      const inferred = epcIssue(c);
      const m = match(units, c.address, c.unit, c.postcode, c.type);
      const reasons = [];
      if (!types.includes(c.type)) reasons.push("certificate type");
      if (!c.issue_date) reasons.push("issue date");
      if (!m.unit || m.score < AUTO_SCORE || m.score - m.second < AUTO_MARGIN) reasons.push("which property");
      if (c.confidence === "low") reasons.push("hard to read");
      const item = { ...c, file: file.name, blob: file, inferred, suggested: m.unit && m.score > 20 ? m.unit.name : "" };
      if (!reasons.length && existing.has(`${m.unit.name}|${c.type}|${serial(c.issue_date)}`)) {
        finish(sub, "skipped", `Already in the tracker: ${m.unit.name} · ${c.type} · ${ukDate(c.issue_date)}`);
      } else if (reasons.length) {
        reviewForm(sub, item, units, types, "Please check: " + reasons.join(", ") + ".");
      } else {
        const n = await save(item, m.unit.name);
        if (!n) finish(sub, "skipped", `Already in the tracker: ${m.unit.name} · ${c.type} · ${ukDate(c.issue_date)}`);
        else finish(sub, "saved", `Saved: ${m.unit.name} · ${c.type} · ${ukDate(c.issue_date)}${c.expiry_or_next_due ? " → " + ukDate(c.expiry_or_next_due) : ""}`);
      }
    }
  } catch (e) {
    finish(card, "error", e.message || String(e));
  }
}

/* copy the original into Certificates/<Building>/<Unit>/ on this Mac, so old certificates are kept */
async function archive(c, propertyName) {
  const enc = encodeURIComponent;
  const res = await fetch(`${ARCHIVE}/archive`, { method: "POST", body: c.blob, headers: { "Content-Type": "application/octet-stream",
    "X-Filename": enc(c.file), "X-Property": enc(propertyName), "X-Type": enc(c.type), "X-Issue": enc(c.issue_date || "") } });
  if (!res.ok) throw new Error((await res.json().catch(() => ({}))).message || `HTTP ${res.status}`);
  return res.json();
}

async function checkArchive() {
  try { archiveOn = (await fetch(`${ARCHIVE}/ping`)).ok; } catch (e) { archiveOn = false; }
  showSummary();
}

async function save(c, propertyName) {
  let file = c.file, link = "", warn = "";
  try {
    const a = await archive(c, propertyName);
    file = a.name; link = a.file; archiveOn = true;
  } catch (e) {
    archiveOn = false; warn = "PDF not archived (archive service off)";
  }
  const notes = [warn, c.inferred, c.engineer && `Engineer: ${c.engineer}`, c.remedial_work && `Remedial: ${c.remedial_work}`].filter(Boolean).join("; ");
  return writeCert({ property: propertyName, type: c.type, issue: c.issue_date, expiry: c.expiry_or_next_due, result: c.result,
    company: c.company, number: c.certificate_number, file, link, notes });
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
    `Saved ${stats.saved} · To check ${stats.review} · Skipped ${stats.skipped} · Errors ${stats.error}` +
    (archiveOn ? " · Archive on" : " · ⚠ Archive off: PDFs are not being copied");
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
      const n = await save(item, prop);
      stats.review--;
      finish(card, n ? "saved" : "skipped", `${n ? "Saved" : "Already in the tracker"}: ${prop} · ${item.type} · ${ukDate(item.issue_date)}`);
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
  checkArchive();
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
