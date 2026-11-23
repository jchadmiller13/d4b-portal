/* TFD District 4-B · Upload roster (reads the District 4-B Master Roster workbook) */
(function(){
const RU = { changes:null };
const LABEL = { name:"Name", first:"First name", last:"Last name", rank:"Rank", unit:"Unit", station:"Company", shift:"Shift", csGroup:"City Shift Group", roic:"ROIC #", empId:"Employee #", email:"Email", phone:"Phone", stationNo:"Station #", hireDate:"Hire date", licensure:"Licensure", licYear:"License year", refresher:"EMS refresher", protocol:"Protocol", physical:"Physical", physicalSched:"Physical scheduled", physicalTime:"Physical time", mandatory:"Mandatory", pfa:"PFA", pat:"PAT", ppe:"PPE", hazmat:"HazMat", doTraining:"DO training", q1:"Q1 hours", q2:"Q2 hours", q3:"Q3 hours", q4:"Q4 hours", qNeed:"Hours needed" };

function pad(n){ return String(n).padStart(2,"0"); }
function cv(v){
  if (v === null || v === undefined) return "";
  if (v instanceof Date){ if (isNaN(v)) return ""; const d = new Date(v.getTime() + 43200000); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }
  let s = String(v).replace(/ /g," ").trim();
  if (["0","#REF!","n/a","N/A"].includes(s)) return "";
  if (/^\d+\.0$/.test(s)) s = s.slice(0,-2);
  return s;
}
function roicOf(v){ const s = cv(v); return /^\d+$/.test(s) ? s.padStart(6,"0") : s; }
function num(v){ if (v === null || v === undefined || v === "" ) return null; const n = Number(v); return isNaN(n) ? null : n; }
const UNIT = { E:"Engine", L:"Ladder", C:"Car" };
function company(u){ const m = String(u).match(/^([ELC])0*(\d+)/); return m ? `${UNIT[m[1]]} ${m[2]}` : u; }

function findSheet(sheets, name){
  if (sheets[name]) return sheets[name];
  const k = Object.keys(sheets).find(n => n.trim().toLowerCase() === name.trim().toLowerCase());
  return k ? sheets[k] : null;
}
function body(sheets, name){ const r = findSheet(sheets, name) || []; return r.slice(1).filter(x => x && x.some(v => v !== null && v !== undefined && v !== "")); }
function idx(sheets, name, li, fi, ri){
  const byr = {}, bylf = {}, byl = {};
  for (const r of body(sheets, name)){
    const l = cv(r[li]).toLowerCase(), f = cv(r[fi]).toLowerCase();
    if (!l || l.startsWith("#ref")) continue;
    if (ri !== undefined && roicOf(r[ri])) byr[roicOf(r[ri])] = r;
    bylf[l+"|"+f] = r; (byl[l] = byl[l] || []).push(r);
  }
  return (rc, l, f) => { l = l.toLowerCase(); f = f.toLowerCase();
    if (rc && byr[rc]) return byr[rc]; if (bylf[l+"|"+f]) return bylf[l+"|"+f]; if ((byl[l]||[]).length === 1) return byl[l][0]; return null; };
}

/* sheets: { sheetName: [[cell, cell, ...], ...] } including the heading row. Returns member records (no personal details). */
function parseRoster(sheets){
  if (!findSheet(sheets, "D4B COs")) throw new Error("This doesn't look like the District 4-B Master Roster: there's no \"D4B COs\" sheet.");
  const cs = {}; for (const r of body(sheets, "D4B CS")) cs[cv(r[0]).toLowerCase()+"|"+cv(r[1]).toLowerCase()] = r;
  const M = idx(sheets,"Master",5,6,7), INFO = idx(sheets,"Info",0,1,2), EMS = idx(sheets,"EMS Refresher",0,1), PHY = idx(sheets,"Physicals ",0,1,2),
        REC = idx(sheets,"Recordables",0,1,4), T2 = idx(sheets,"T2 - Co Training ",0,1), HIRE = idx(sheets,"Hire Date",0,1,6);
  const out = [];
  for (const r of body(sheets, "D4B COs")){
    const last = cv(r[0]), first = cv(r[1]), unit = cv(r[2]), rank = cv(r[3]);
    if (!last && !first) continue;
    const g = cs[last.toLowerCase()+"|"+first.toLowerCase()];
    const rc = g ? roicOf(g[2]) : "";
    const m = M(rc,last,first), inf = INFO(rc,last,first), e = EMS(rc,last,first), p = PHY(rc,last,first), rec = REC(rc,last,first), t = T2(rc,last,first), h = HIRE(rc,last,first);
    const name = m && cv(m[4]) ? cv(m[4]) : `${first} ${last}`.trim();
    let d = { name, first, last, rank, unit, station: company(unit), shift:"B",
      csGroup: g ? cv(g[3]) : "", roic: rc,
      empId: rec ? cv(rec[5]) : (m ? cv(m[12]) : (inf ? cv(inf[3]) : "")),
      email: m ? cv(m[3]).toLowerCase() : "", phone: m ? cv(m[20]) : (inf ? cv(inf[6]) : ""),
      stationNo: m ? cv(m[9]) : "", hireDate: rec ? cv(rec[6]) : (h ? cv(h[2]) : (m ? cv(m[14]) : "")),
      licensure: e ? cv(e[2]) : (m ? cv(m[15]) : ""), licYear: e ? cv(e[3]) : (m ? cv(m[16]) : ""),
      refresher: e ? cv(e[4]) : "", protocol: e ? cv(e[5]) : "",
      physical: p ? cv(p[4]) : "", physicalSched: p ? cv(p[6]) : "", physicalTime: p && cv(p[7]) ? cv(p[7]).padStart(4,"0") : "",
      mandatory: rec ? cv(rec[7]) : "", pfa: rec ? cv(rec[8]) : "", pat: rec ? cv(rec[9]) : "", ppe: rec ? cv(rec[10]) : "",
      hazmat: rec ? cv(rec[11]) : "", doTraining: rec ? cv(rec[17]) : "" };
    if (t){ [["q1",4],["q2",7],["q3",10],["q4",13]].forEach(([q,ix]) => { const n = num(t[ix]); d[q] = n === null ? 0 : n; }); d.qNeed = num(t[5]) || 48; }
    d = Object.fromEntries(Object.entries(d).filter(([k,v]) => v !== ""));
    out.push(d);
  }
  return out;
}

function workbookRows(wb){
  const sheets = {};
  for (const n of wb.SheetNames){
    const ws = wb.Sheets[n]; if (!ws || !ws["!ref"]) { sheets[n] = []; continue; }
    const rg = XLSX.utils.decode_range(ws["!ref"]); rg.s.r = 0; rg.s.c = 0;
    sheets[n] = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:null, blankrows:true, range:rg });
  }
  return sheets;
}

function matchExisting(rec, used){
  const P = state.personnel.filter(p => !used.has(p.id));
  return (rec.roic && P.find(p => p.roic === rec.roic))
      || P.find(p => (p.name||"").toLowerCase() === rec.name.toLowerCase())
      || (rec.first && rec.last && P.find(p => (p.first||"").toLowerCase() === rec.first.toLowerCase() && (p.last||"").toLowerCase() === rec.last.toLowerCase()))
      || null;
}
function diff(recs){
  const used = new Set(), add = [], upd = [];
  for (const rec of recs){
    const ex = matchExisting(rec, used);
    if (!ex){ add.push(rec); continue; }
    used.add(ex.id);
    const ch = Object.keys(rec).filter(k => String(ex[k] ?? "") !== String(rec[k]));
    if (ch.length) upd.push({ ex, rec, ch });
  }
  const gone = state.personnel.filter(p => !used.has(p.id) && !add.some(a => a === p));
  return { add, upd, gone, total: recs.length };
}

function openRosterUpload(){
  if (typeof XLSX === "undefined"){ toast("The Excel reader didn't load. Reload the page and try again."); return; }
  RU.changes = null;
  showModal(`<h2>Upload roster</h2>
    <div class="hint">Choose your District 4-B Master Roster workbook. You'll see every change before anything is saved. Members are matched by ROIC #, then by name. Addresses, birth dates and emergency contacts in the workbook are never uploaded.</div>
    <label for="rosterFile">Workbook<input id="rosterFile" type="file" accept=".xlsx,.xlsm,.xls"></label>
    <div id="ruBody"></div><div class="err" id="ruErr"></div>
    <div class="mf"><span></span><div class="tools"><button class="btn" data-a="close">Cancel</button><button class="btn pri" data-a="rosterapply" id="ruGo" disabled>Apply changes</button></div></div>`, true);
}

function showDiff(c){
  const row = (n, t) => `<tr><td><b>${esc(n)}</b></td><td>${t}</td></tr>`;
  const updRows = c.upd.map((u,i) => row(u.ex.name, `<label class="inline"><input type="checkbox" id="ru_u${i}" checked> ` + u.ch.map(k => `${esc(LABEL[k]||k)}: <s>${esc(String(u.ex[k] ?? "–"))}</s> → ${esc(String(u.rec[k]))}`).join("<br>") + `</label>`)).join("");
  const addRows = c.add.map((a,i) => row(a.name, `<label class="inline"><input type="checkbox" id="ru_a${i}" checked> New · ${esc([a.rank,a.station,a.csGroup?`City Shift Group ${a.csGroup}`:""].filter(Boolean).join(" · "))}</label>`)).join("");
  const goneRows = c.gone.map(g => row(g.name, `<span class="hint">Not in the workbook. Left on the roster; remove them on the Roster page if they've left the district.</span>`)).join("");
  const none = !c.upd.length && !c.add.length;
  $("#ruBody").innerHTML = `<h3 style="margin-top:6px">${c.total} members in the workbook · ${c.upd.length} with changes · ${c.add.length} new</h3>
    ${none ? `<div class="hint">The roster already matches the workbook. Nothing to change.</div>` : `<div class="hint">Untick anything you don't want. A change you made in the app (like a move between companies) shows here if the workbook still has the old value; untick it to keep the app's version.</div>`}
    <div class="tw"><table><tbody>${updRows}${addRows}${goneRows}</tbody></table></div>`;
  $("#ruGo").disabled = none;
}

function readRoster(file){
  const r = new FileReader();
  r.onload = () => {
    try {
      const wb = XLSX.read(new Uint8Array(r.result), { type:"array", cellDates:true });
      RU.changes = diff(parseRoster(workbookRows(wb)));
      $("#ruErr").textContent = ""; showDiff(RU.changes);
    } catch(e){ $("#ruErr").textContent = e && e.message && /Master Roster/.test(e.message) ? e.message : "That file couldn't be read as the roster workbook."; $("#ruGo").disabled = true; }
  };
  r.readAsArrayBuffer(file);
}

async function applyRoster(btn){
  const c = RU.changes; if (!c) return;
  if (!db){ $("#ruErr").textContent = writeErr(); return; }
  btn.disabled = true; let n = 0, u = 0;
  try {
    for (let i = 0; i < c.upd.length; i++){
      const box = $("#ru_u"+i); if (box && !box.checked) continue;
      const { ex, rec, ch } = c.upd[i]; const patch = {}; ch.forEach(k => patch[k] = rec[k]);
      await db.collection("personnel").doc(ex.id).update(patch); u++;
      btn.textContent = `Saving… ${u+n}`;
    }
    for (let i = 0; i < c.add.length; i++){
      const box = $("#ru_a"+i); if (box && !box.checked) continue;
      await db.collection("personnel").doc().set(c.add[i]); n++;
      btn.textContent = `Saving… ${u+n}`;
    }
  } catch(e){ $("#ruErr").textContent = writeErr(e) + ` (${u+n} saved before the error)`; btn.disabled = false; return; }
  closeModal(); RU.changes = null;
  toast(`Roster updated: ${u} changed, ${n} added`);
}

document.addEventListener("click", e => {
  const el = e.target.closest("[data-a]"); if (!el) return;
  if (el.dataset.a === "rosterup") openRosterUpload();
  if (el.dataset.a === "rosterapply") applyRoster(el);
});
document.addEventListener("change", e => { if (e.target.id === "rosterFile" && e.target.files[0]) readRoster(e.target.files[0]); });

window.TFDRoster = { parseRoster, diff, cv };
})();
