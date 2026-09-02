import { Streamlit } from "https://cdn.jsdelivr.net/npm/streamlit-component-lib@2.0.0/+esm";

const root = document.getElementById("app");
let lastRequestId = null;
let running = false;
let generator = null;

function esc(s) { return String(s ?? "").replace(/[&<>\"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }
function paint(title, detail, progress = null, kind = "info") {
  const pct = progress == null ? "" : `<div style="height:6px;background:#e5e7eb;border-radius:99px;overflow:hidden;margin-top:8px"><div style="height:100%;width:${Math.max(2,Math.min(100,progress))}%;background:#2563eb"></div></div>`;
  const bg = kind === "error" ? "#fef2f2" : kind === "ok" ? "#f0fdf4" : "#f8fafc";
  const border = kind === "error" ? "#fecaca" : kind === "ok" ? "#bbf7d0" : "#e2e8f0";
  root.innerHTML = `<div style="font-family:system-ui,-apple-system,Segoe UI,sans-serif;padding:12px 14px;border:1px solid ${border};border-radius:10px;background:${bg};font-size:14px;color:#0f172a"><div style="font-weight:650">${esc(title)}</div><div style="margin-top:4px;color:#475569">${esc(detail)}</div>${pct}</div>`;
  Streamlit.setFrameHeight(92);
}
function extractJson(text) {
  const raw = String(text || "").trim();
  try { return JSON.parse(raw); } catch (_) {}
  const start = raw.indexOf("{"); if (start < 0) throw new Error("Model JSON üretmedi.");
  let depth=0,inString=false,escaped=false;
  for(let i=start;i<raw.length;i++) { const ch=raw[i]; if(inString){if(escaped)escaped=false;else if(ch==="\\")escaped=true;else if(ch==='"')inString=false;continue;} if(ch==='"')inString=true; else if(ch==="{")depth++; else if(ch==="}"){depth--;if(depth===0)return JSON.parse(raw.slice(start,i+1));} }
  throw new Error("Model yanıtındaki JSON tamamlanamadı.");
}
async function ensureGenerator(modelId) {
  if (generator) return generator;
  if (!("gpu" in navigator)) throw new Error("WebGPU bulunamadı. Güncel Chrome/Edge ve donanım hızlandırmayı kullanın.");
  paint("Tarayıcı AI hazırlanıyor", "Model ilk kullanımda indirilir; sonra tarayıcı önbelleğinden açılır.", 8);
  const mod = await import("https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.8.1/+esm");
  mod.env.allowLocalModels = false;
  generator = await mod.pipeline("text-generation", modelId, {dtype:"q4",device:"webgpu",progress_callback:(p)=>{const v=typeof p?.progress==="number"?Math.round(p.progress):18;paint("Tarayıcı AI hazırlanıyor",p?.file?`Model dosyası: ${p.file}`:"Model yükleniyor...",Math.max(8,Math.min(70,v*.7)));}});
  return generator;
}
async function run(args) {
  running=true;
  try {
    const modelId=args.model_id||"onnx-community/Qwen2.5-0.5B-Instruct";
    const gen=await ensureGenerator(modelId);
    paint("Belge bilgileri AI ile doğrulanıyor","İşlem bu cihazın tarayıcısında yapılıyor; belge metni harici AI API'sine gönderilmiyor.",78);
    const messages=[{role:"system",content:"You are a precise data extraction engine. Read Turkish patent application source text. Never invent data. Return only valid JSON."},{role:"user",content:args.prompt||""}];
    const output=await gen(messages,{max_new_tokens:Number(args.max_new_tokens||650),do_sample:false,repetition_penalty:1.02,return_full_text:true});
    let content=output?.[0]?.generated_text;
    if(Array.isArray(content)) content=content.at(-1)?.content||"";
    else if(typeof content!=="string") content="";
    const parsed=extractJson(content);
    paint("Tarayıcı AI tamamlandı","Değerler kaynak metinle sunucuda yeniden doğrulanıyor.",100,"ok");
    Streamlit.setComponentValue({request_id:args.request_id,ok:true,data:parsed,model:modelId});
  } catch(e) {
    const msg=e?.message||String(e); paint("Tarayıcı AI çalıştırılamadı",msg,null,"error"); Streamlit.setComponentValue({request_id:args.request_id,ok:false,error:msg});
  } finally { running=false; }
}
function onRender(event) { const args=event.detail.args||{}; Streamlit.setFrameHeight(92); if(!args.request_id){paint("Tarayıcı AI hazır","Belge analizi başlatıldığında yalnız bu tarayıcıda çalışır.");return;} if(args.request_id!==lastRequestId&&!running){lastRequestId=args.request_id;run(args);} }
Streamlit.events.addEventListener(Streamlit.RENDER_EVENT,onRender); Streamlit.setComponentReady(); Streamlit.setFrameHeight(92); paint("Tarayıcı AI hazırlanıyor","Analiz isteği bekleniyor...");
