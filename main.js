import { Streamlit } from "https://cdn.jsdelivr.net/npm/streamlit-component-lib@2.0.0/+esm";

const root = document.getElementById("app");
let lastRequestId = null;
let running = false;
let ner = null;

function esc(s) {
  return String(s ?? "").replace(/[&<>\"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
}
function paint(title, detail, progress = null, kind = "info") {
  const pct = progress == null ? "" : `<div style="height:6px;background:#e5e7eb;border-radius:99px;overflow:hidden;margin-top:8px"><div style="height:100%;width:${Math.max(2,Math.min(100,progress))}%;background:#2563eb"></div></div>`;
  const bg = kind === "error" ? "#fef2f2" : kind === "ok" ? "#f0fdf4" : "#f8fafc";
  const border = kind === "error" ? "#fecaca" : kind === "ok" ? "#bbf7d0" : "#e2e8f0";
  root.innerHTML = `<div style="font-family:system-ui,-apple-system,Segoe UI,sans-serif;padding:12px 14px;border:1px solid ${border};border-radius:10px;background:${bg};font-size:14px;color:#0f172a"><div style="font-weight:650">${esc(title)}</div><div style="margin-top:4px;color:#475569">${esc(detail)}</div>${pct}</div>`;
  Streamlit.setFrameHeight(92);
}
function withTimeout(promise, ms, message) {
  let timer;
  const timeout = new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(message)), ms); });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}
function plain(s) {
  return String(s || "").toLocaleLowerCase("tr-TR")
    .normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
    .replace(/ı/g,"i").replace(/ş/g,"s").replace(/ğ/g,"g").replace(/ü/g,"u").replace(/ö/g,"o").replace(/ç/g,"c");
}
const focusNeedles = [
  "hak sahibi","başvuru sahibi","başvuru sahib","buluş sahibi","buluşçu","mucit","buluşu yapan",
  "unvan","ad soyad","adı soyadı","tckn","tc kimlik","vkn","vergi","adres","uyruk","ülke","ilçe",
  "telefon","e-posta","eposta","email","doğum","gizlensin","erken yayın","erken yayım","tübitak","kosgeb","proje"
].map(plain);
function focusText(raw) {
  const lines = String(raw || "").split(/\r?\n/).map(x => x.trim()).filter(Boolean);
  if (!lines.length) return "";
  const keep = new Set();
  for (let i=0;i<lines.length;i++) {
    const n = plain(lines[i]);
    const relevant = focusNeedles.some(k => n.includes(k)) || /@/.test(lines[i]) || /\b\d{10,11}\b/.test(lines[i]) || /anonim\s+şirket|limited\s+şirket|a\.?\s*ş\.?|ltd\.?/i.test(lines[i]);
    if (relevant) for (let j=Math.max(0,i-3); j<Math.min(lines.length,i+5); j++) keep.add(j);
  }
  if (!keep.size) return lines.slice(0,55).join("\n").slice(0,5000);
  return [...keep].sort((a,b)=>a-b).map(i=>lines[i]).join("\n").slice(0,6000);
}
function chunkText(text, maxChars=900) {
  const lines = String(text || "").split(/\r?\n/);
  const chunks = [];
  let buf = [];
  let len = 0;
  for (const line of lines) {
    const add = line.length + 1;
    if (buf.length && len + add > maxChars) {
      chunks.push(buf.join("\n"));
      buf = buf.slice(-2); // iki satır bağlam üst üste kalsın
      len = buf.reduce((a,x)=>a+x.length+1,0);
    }
    buf.push(line);
    len += add;
  }
  if (buf.length) chunks.push(buf.join("\n"));
  return chunks.filter(x=>x.trim());
}
function normalizeLabel(x) {
  let s = String(x || "").toUpperCase();
  s = s.replace(/^[BI]-/, "");
  if (s.includes("PER")) return "PER";
  if (s.includes("ORG")) return "ORG";
  if (s.includes("LOC")) return "LOC";
  return s;
}
function dedupeEntities(rows) {
  const out=[]; const seen=new Set();
  for (const r of rows) {
    const key = `${r.source}|${r.label}|${plain(r.word)}|${plain(r.before).slice(-80)}|${plain(r.after).slice(0,80)}`;
    if (!r.word || seen.has(key)) continue;
    seen.add(key); out.push(r);
  }
  return out;
}
async function ensureNer(modelId) {
  if (ner) return ner;
  paint("CPU AI hazırlanıyor", "İlk kullanımda yaklaşık 135 MB çok dilli NER modeli indirilir; GPU gerekmez.", 5);
  const mod = await import("https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.8.1/+esm");
  mod.env.allowLocalModels = false;
  // En geniş tarayıcı uyumluluğu için WASM tek iş parçacığı. WebGPU kullanılmaz.
  try { mod.env.backends.onnx.wasm.numThreads = 1; } catch (_) {}
  ner = await withTimeout(
    mod.pipeline("token-classification", modelId, {
      dtype: "q8",
      device: "wasm",
      progress_callback: (p) => {
        const raw = typeof p?.progress === "number" ? p.progress : 0;
        const v = raw > 1 ? raw : raw * 100;
        const pct = Math.max(5, Math.min(55, 5 + Math.round(v * 0.5)));
        paint("CPU AI hazırlanıyor", p?.file ? `Model dosyası: ${p.file}` : "Model indiriliyor / önbellekten açılıyor...", pct);
      },
    }),
    180000,
    "CPU AI modeli 3 dakika içinde hazırlanamadı. İnternet bağlantısını kontrol edip tekrar deneyin."
  );
  return ner;
}
async function analyzeSource(pipe, source, text, startIndex, totalChunks) {
  const focused = focusText(text);
  const chunks = chunkText(focused);
  const rows = [];
  let idx = startIndex;
  for (const chunk of chunks) {
    idx += 1;
    const pct = 58 + Math.round((idx / Math.max(1,totalChunks)) * 36);
    paint("Belge bilgileri CPU AI ile tanınıyor", `${source}: kişi / kurum adları belirleniyor (${idx}/${totalChunks})`, Math.min(94,pct));
    const output = await pipe(chunk, {aggregation_strategy:"simple"});
    for (const e of (output || [])) {
      const label = normalizeLabel(e.entity_group || e.entity || e.label);
      if (!["PER","ORG","LOC"].includes(label)) continue;
      const score = Number(e.score || 0);
      if (score < 0.55) continue;
      const word = String(e.word || "").replace(/\s*##\s*/g, "").trim();
      const start = Number.isFinite(e.start) ? Number(e.start) : Math.max(0, chunk.indexOf(word));
      const end = Number.isFinite(e.end) ? Number(e.end) : start + word.length;
      rows.push({
        source,
        label,
        word,
        score,
        before: chunk.slice(Math.max(0,start-650), Math.max(0,start)),
        after: chunk.slice(Math.max(0,end), Math.min(chunk.length,end+900)),
      });
    }
  }
  return {rows, consumed: chunks.length};
}
async function run(args) {
  running = true;
  try {
    const modelId = args.model_id || "Xenova/distilbert-base-multilingual-cased-ner-hrl";
    const sources = Array.isArray(args.sources) ? args.sources.filter(x => x && x.source && x.text) : [];
    if (!sources.length) throw new Error("CPU AI için okunabilir başvuru bilgi kaynağı bulunamadı.");
    const pipe = await ensureNer(modelId);
    const prepared = sources.map(s => ({source:String(s.source), text:focusText(String(s.text||""))}));
    const totalChunks = prepared.reduce((n,s)=>n+chunkText(s.text).length,0);
    let consumed=0; let entities=[];
    for (const s of prepared) {
      const res = await withTimeout(
        analyzeSource(pipe, s.source, s.text, consumed, totalChunks),
        Number(args.inference_timeout_ms || 120000),
        "CPU AI belge analizi zaman aşımına uğradı. Kurallı sonuç kullanılacak."
      );
      consumed += res.consumed;
      entities.push(...res.rows);
    }
    entities = dedupeEntities(entities);
    paint("CPU AI tamamlandı", `${entities.length} kişi/kurum/konum adayı bulundu; kaynak ilişkileri sunucuda doğrulanıyor.`, 100, "ok");
    Streamlit.setComponentValue({request_id:args.request_id, ok:true, data:{entities}, model:`${modelId} (CPU/WASM q8)`});
  } catch (e) {
    const msg = e?.message || String(e);
    paint("CPU AI çalıştırılamadı", msg, null, "error");
    Streamlit.setComponentValue({request_id:args.request_id, ok:false, error:msg});
  } finally { running=false; }
}
function onRender(event) {
  const args = event.detail.args || {};
  Streamlit.setFrameHeight(92);
  if (!args.request_id) { paint("CPU AI hazır", "Belge analizi başlatıldığında bu tarayıcının CPU'sunda çalışır."); return; }
  if (args.request_id !== lastRequestId && !running) { lastRequestId=args.request_id; run(args); }
}
Streamlit.events.addEventListener(Streamlit.RENDER_EVENT, onRender);
Streamlit.setComponentReady();
Streamlit.setFrameHeight(92);
paint("CPU AI hazırlanıyor", "Analiz isteği bekleniyor...");
