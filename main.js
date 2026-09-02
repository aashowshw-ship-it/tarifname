import { Streamlit } from "https://cdn.jsdelivr.net/npm/streamlit-component-lib@2.0.0/+esm";

const root = document.getElementById("app");
let lastRequestId = null;
let running = false;
let classifier = null;

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

const needles = [
  "hak sahibi","başvuru sahibi","başvuru sahib","buluş sahibi","buluşçu","mucit","buluşu yapan",
  "unvan","ad soyad","adı soyadı","tckn","tc kimlik","vkn","vergi","adres","uyruk","ülke","ilçe",
  "telefon","e-posta","eposta","email","doğum","gizlensin","erken yayın","erken yayım","tübitak","kosgeb","proje",
  "yetkili","imza","temsilci","irtibat"
].map(plain);

function relevantLine(line) {
  const n = plain(line);
  return needles.some(k => n.includes(k)) || /@/.test(line) || /\b\d{10,11}\b/.test(line) ||
    /anonim\s+şirket|limited\s+şirket|a\.?\s*ş\.?|ltd\.?/i.test(line) || /\b(?:EVET|HAYIR)\b/i.test(line) || /\t/.test(line);
}

function makeBlocks(raw, maxBlocks = 16) {
  const lines = String(raw || "").replace(/\r/g, "").split("\n");
  const picks = [];
  const seen = new Set();
  const add = (start, end) => {
    const text = lines.slice(Math.max(0,start), Math.min(lines.length,end)).map(x=>x.trimEnd()).filter(x=>x.trim()).join("\n").trim();
    if (!text || text.length < 12) return;
    const clipped = text.slice(0, 1500);
    const key = plain(clipped).replace(/\s+/g," ").slice(0,1200);
    if (!key || seen.has(key)) return;
    seen.add(key); picks.push(clipped);
  };
  for (let i=0;i<lines.length && picks.length<maxBlocks;i++) {
    if (!relevantLine(lines[i])) continue;
    add(i-2, i+4);
  }
  // Hiç aday bulunamazsa belgenin başından birkaç kısa pencere ver; AI rolü kendisi seçsin.
  if (!picks.length) {
    for (let i=0;i<lines.length && picks.length<4;i+=6) add(i, i+7);
  }
  return picks.slice(0,maxBlocks);
}

const ROLE_LABELS = [
  "hak sahibi veya başvuru sahibi bilgileri",
  "buluş sahibi, buluşçu veya mucit bilgileri",
  "başvuru tercihleri veya beyan cevapları",
  "yetkili, imza sahibi veya yalnız iletişim kişisi bilgileri",
  "diğer bilgi"
];
const ROLE_KEYS = {
  "hak sahibi veya başvuru sahibi bilgileri": "applicant",
  "buluş sahibi, buluşçu veya mucit bilgileri": "inventor",
  "başvuru tercihleri veya beyan cevapları": "options",
  "yetkili, imza sahibi veya yalnız iletişim kişisi bilgileri": "contact",
  "diğer bilgi": "other",
};

async function ensureClassifier(modelId) {
  if (classifier) return classifier;
  paint("CPU AI hazırlanıyor", "Türkçe dahil çok dilli anlam sınıflandırma modeli ilk kullanımda tarayıcıya indirilir. GPU gerekmez.", 5);
  const mod = await import("https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.8.1/+esm");
  mod.env.allowLocalModels = false;
  try { mod.env.backends.onnx.wasm.numThreads = 1; } catch (_) {}
  classifier = await withTimeout(
    mod.pipeline("zero-shot-classification", modelId, {
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
  return classifier;
}

async function classifyBlock(pipe, source, text, index, total) {
  const pct = 58 + Math.round((index / Math.max(1,total)) * 36);
  paint("Belge yapısı CPU AI ile yorumlanıyor", `${source}: bilgi bloğunun rolü belirleniyor (${index}/${total})`, Math.min(94,pct));
  const out = await pipe(text, ROLE_LABELS, {
    multi_label: false,
    hypothesis_template: "Bu metin {} içeriyor."
  });
  const labels = Array.isArray(out?.labels) ? out.labels : [];
  const scores = Array.isArray(out?.scores) ? out.scores.map(Number) : [];
  if (!labels.length) return null;
  const role = ROLE_KEYS[labels[0]] || "other";
  const score = Number(scores[0] || 0);
  const second = Number(scores[1] || 0);
  return {source, text, role, score, margin: Math.max(0, score-second), labels, scores};
}

async function run(args) {
  running = true;
  try {
    const modelId = args.model_id || "onnx-community/multilingual-MiniLMv2-L6-mnli-xnli-ONNX";
    const sources = Array.isArray(args.sources) ? args.sources.filter(x => x && x.source && x.text) : [];
    if (!sources.length) throw new Error("CPU AI için okunabilir başvuru bilgi kaynağı bulunamadı.");
    const pipe = await ensureClassifier(modelId);
    const prepared = [];
    for (const s of sources) {
      for (const block of makeBlocks(String(s.text || ""), Number(args.max_blocks_per_source || 16))) {
        prepared.push({source:String(s.source), text:block});
      }
    }
    if (!prepared.length) throw new Error("Anlamlandırılabilecek bilgi bloğu bulunamadı.");
    const blocks = [];
    for (let i=0;i<prepared.length;i++) {
      const r = await withTimeout(
        classifyBlock(pipe, prepared[i].source, prepared[i].text, i+1, prepared.length),
        Number(args.inference_timeout_ms || 90000),
        "CPU AI belge analizi zaman aşımına uğradı. Kurallı sonuç kullanılacak."
      );
      if (!r) continue;
      // Contact/other bloklarını veri üretmek için kullanma. Rol için düşük güveni de dışla.
      const threshold = r.role === "options" ? 0.40 : 0.46;
      if (["applicant","inventor","options"].includes(r.role) && r.score >= threshold && (r.margin >= 0.015 || r.score >= 0.62)) {
        blocks.push(r);
      }
    }
    paint("CPU AI tamamlandı", `${blocks.length} güvenilir başvuru bilgi bloğu rolüne göre ayrıldı; değerler kaynak metinde doğrulanıyor.`, 100, "ok");
    Streamlit.setComponentValue({request_id:args.request_id, ok:true, data:{blocks}, model:`${modelId} (CPU/WASM q8)`});
  } catch (e) {
    const msg = e?.message || String(e);
    paint("CPU AI çalıştırılamadı", msg, null, "error");
    Streamlit.setComponentValue({request_id:args.request_id, ok:false, error:msg});
  } finally { running=false; }
}

function onRender(event) {
  const args = event.detail.args || {};
  Streamlit.setFrameHeight(92);
  if (!args.request_id) { paint("CPU AI hazır", "Belge analizi başlatıldığında form düzenini bu tarayıcının CPU'sunda anlamlandırır."); return; }
  if (args.request_id !== lastRequestId && !running) { lastRequestId=args.request_id; run(args); }
}
Streamlit.events.addEventListener(Streamlit.RENDER_EVENT, onRender);
Streamlit.setComponentReady();
Streamlit.setFrameHeight(92);
paint("CPU AI hazırlanıyor", "Analiz isteği bekleniyor...");
