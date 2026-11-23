/* TFD District 4-B · Dashboard weather (Open-Meteo forecast + National Weather Service alerts, no key needed) */
(function(){
const LOC = { name:"Tulsa", lat:36.154, lon:-95.993 };
const WX = { data:null, alerts:[], at:0, err:"", busy:false };
const CODES = { 0:"Clear",1:"Mostly clear",2:"Partly cloudy",3:"Overcast",45:"Fog",48:"Freezing fog",51:"Light drizzle",53:"Drizzle",55:"Heavy drizzle",56:"Freezing drizzle",57:"Freezing drizzle",61:"Light rain",63:"Rain",65:"Heavy rain",66:"Freezing rain",67:"Freezing rain",71:"Light snow",73:"Snow",75:"Heavy snow",77:"Snow grains",80:"Rain showers",81:"Rain showers",82:"Heavy showers",85:"Snow showers",86:"Heavy snow showers",95:"Thunderstorms",96:"Storms with hail",99:"Storms with hail" };
const ICON = c => c === 0 || c === 1 ? "☀️" : c === 2 ? "⛅" : c === 3 ? "☁️" : c === 45 || c === 48 ? "🌫️" : (c >= 51 && c <= 67) || (c >= 80 && c <= 82) ? "🌧️" : (c >= 71 && c <= 77) || c === 85 || c === 86 ? "❄️" : c >= 95 ? "⛈️" : "🌡️";
const r0 = v => v == null || isNaN(v) ? "–" : Math.round(v);
const hr = iso => { const h = +iso.slice(11,13); return (h % 12 || 12) + (h < 12 ? "a" : "p"); };
const tm = iso => { if (!iso) return "–"; const h = +iso.slice(11,13), m = iso.slice(14,16); return `${h % 12 || 12}:${m} ${h < 12 ? "AM" : "PM"}`; };

async function load(force){
  if (WX.busy || (!force && WX.data && Date.now() - WX.at < 15*60*1000)) return;
  WX.busy = true;
  try {
    const u = `https://api.open-meteo.com/v1/forecast?latitude=${LOC.lat}&longitude=${LOC.lon}&current=temperature_2m,apparent_temperature,relative_humidity_2m,weather_code,wind_speed_10m,wind_gusts_10m,wind_direction_10m&hourly=temperature_2m,precipitation_probability,weather_code&daily=weather_code,temperature_2m_max,temperature_2m_min,apparent_temperature_max,apparent_temperature_min,precipitation_probability_max,precipitation_sum,wind_speed_10m_max,wind_gusts_10m_max,sunrise,sunset&temperature_unit=fahrenheit&wind_speed_unit=mph&precipitation_unit=inch&timezone=America%2FChicago&forecast_days=8`;
    const r = await fetch(u); if (!r.ok) throw new Error(r.status);
    WX.data = await r.json(); WX.at = Date.now(); WX.err = "";
  } catch(e){ WX.err = "The forecast couldn't be loaded right now."; }
  try {
    const a = await fetch(`https://api.weather.gov/alerts/active?point=${LOC.lat},${LOC.lon}`, { headers:{ Accept:"application/geo+json" } });
    if (a.ok){ const j = await a.json(); WX.alerts = (j.features||[]).map(f => f.properties).filter(p => p && p.event); }
  } catch(e){}
  WX.busy = false;
  paint();
}

function dirName(deg){ if (deg == null) return ""; return ["N","NE","E","SE","S","SW","W","NW"][Math.round(deg/45) % 8]; }

function bodyHtml(day){
  const d = WX.data;
  if (!d) return `<div class="hint">${WX.err ? esc(WX.err) + ` <button class="btn sm" data-a="wxretry">Try again</button>` : "Loading the forecast…"}</div>`;
  const i = (d.daily.time||[]).indexOf(day);
  if (i < 0) return `<div class="hint">No forecast is available yet for ${esc(day)}.</div>`;
  const D = k => d.daily[k][i];
  const code = D("weather_code");
  const isToday = day === today();
  const cur = d.current || {};
  // shift hours: 7 AM on the shift day to 7 AM the next morning, every 3 hours
  const H = d.hourly, start = H.time.indexOf(`${day}T07:00`);
  const hours = start < 0 ? [] : [0,3,6,9,12,15,18,21,24].map(o => start + o).filter(j => j < H.time.length);
  const alerts = WX.alerts.length ? `<div class="wxal">${WX.alerts.map(p => `<div><b>⚠️ ${esc(p.event)}</b>${p.ends || p.expires ? ` <span>until ${esc(new Date(p.ends||p.expires).toLocaleString("en-US",{ weekday:"short", hour:"numeric", minute:"2-digit", timeZone:"America/Chicago" }))}</span>` : ""}${p.headline ? `<div class="hint">${esc(p.headline)}</div>` : ""}</div>`).join("")}</div>` : "";
  return `${alerts}
  <div class="wxmain">
    <div class="wxbig"><span class="wxi" aria-hidden="true">${ICON(code)}</span><div><div class="wxt">${r0(D("temperature_2m_max"))}° <span>/ ${r0(D("temperature_2m_min"))}°</span></div><div class="wxc">${esc(CODES[code] || "")}</div></div></div>
    <div class="wxstats">
      ${isToday && cur.temperature_2m != null ? `<div><span>Now</span><b>${r0(cur.temperature_2m)}°</b> feels ${r0(cur.apparent_temperature)}°</div>` : ""}
      <div><span>Feels like</span><b>${r0(D("apparent_temperature_max"))}°</b> / ${r0(D("apparent_temperature_min"))}°</div>
      <div><span>Rain chance</span><b>${r0(D("precipitation_probability_max"))}%</b>${D("precipitation_sum") ? ` · ${(+D("precipitation_sum")).toFixed(2)} in` : ""}</div>
      <div><span>Wind</span><b>${r0(D("wind_speed_10m_max"))} mph</b> gusts ${r0(D("wind_gusts_10m_max"))}${isToday && cur.wind_direction_10m != null ? ` · from ${dirName(cur.wind_direction_10m)}` : ""}</div>
      ${isToday && cur.relative_humidity_2m != null ? `<div><span>Humidity</span><b>${r0(cur.relative_humidity_2m)}%</b></div>` : ""}
      <div><span>Sun</span>${tm(D("sunrise"))} – ${tm(D("sunset"))}</div>
    </div>
  </div>
  ${hours.length ? `<div class="wxhrs" aria-label="Shift forecast, 7 AM to 7 AM">${hours.map(j => `<div><span>${hr(H.time[j])}</span><i aria-hidden="true">${ICON(H.weather_code[j])}</i><b>${r0(H.temperature_2m[j])}°</b><small>${r0(H.precipitation_probability[j])}%</small></div>`).join("")}</div>` : ""}
  <div class="hint" style="margin-top:6px">${esc(LOC.name)} · Forecast from Open-Meteo, alerts from the National Weather Service · updated ${new Date(WX.at).toLocaleTimeString("en-US",{ hour:"numeric", minute:"2-digit" })}</div>`;
}

let curDay = null;
function paint(){ const el = document.getElementById("wxBody"); if (el && curDay) el.innerHTML = bodyHtml(curDay); }

window.wxPanel = function(day){
  curDay = day;
  load(false);
  return `<section class="panel"><div class="ph"><h3>Weather · ${esc(dowOf(day))} shift</h3></div><div id="wxBody">${bodyHtml(day)}</div></section>`;
};

document.addEventListener("click", e => { const el = e.target.closest("[data-a='wxretry']"); if (el) load(true); });
setInterval(() => { if (document.getElementById("wxBody")) load(false); }, 5*60*1000);

const css = document.createElement("style");
css.textContent = `
.wxmain{display:flex;flex-wrap:wrap;gap:12px 24px;align-items:center}
.wxbig{display:flex;align-items:center;gap:10px}
.wxi{font-size:40px;line-height:1}
.wxt{font-family:var(--display);font-size:34px;font-weight:700;line-height:1}
.wxt span{font-size:22px;color:var(--muted)}
.wxc{color:var(--muted);font-size:14px}
.wxstats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:4px 16px;flex:1;min-width:240px;font-size:14px}
.wxstats span{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted)}
.wxhrs{display:flex;gap:4px;overflow-x:auto;margin-top:10px;padding:6px;border:1px solid var(--line);border-radius:6px;background:var(--soft)}
.wxhrs>div{flex:1 0 52px;text-align:center;font-size:12px;display:flex;flex-direction:column;align-items:center}
.wxhrs span{color:var(--muted)} .wxhrs i{font-style:normal;font-size:18px} .wxhrs small{color:var(--info,#1a5fb4)}
.wxal{margin-bottom:10px;padding:8px 10px;border-radius:6px;background:#fdecea;border:1px solid #b3261e;color:#5c1410;font-size:14px}
.wxal>div+div{margin-top:6px}
.wxal span{color:#5c1410}`;
document.head.appendChild(css);
})();
