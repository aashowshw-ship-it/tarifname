from __future__ import annotations

import json
import os
from pathlib import Path

import pandas as pd
import streamlit as st

from modules.ai import ask_json
from modules.docgen import build_docx
from modules.parsers import extract_text
from modules.prompts import drafting_prompt, extraction_prompt, literature_prompt
from modules.validators import validate_draft

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = BASE_DIR / "assets" / "Tarifname_181176_template.docx"

st.set_page_config(page_title="Tarifname Atölyesi", page_icon="📄", layout="wide")

st.markdown(
    """
    <style>
      .block-container {max-width: 1250px; padding-top: 1.5rem;}
      .hero {padding: 1.4rem 1.6rem; background: white; border: 1px solid #e7e2df; border-radius: 16px; margin-bottom: 1rem;}
      .hero h1 {margin: 0 0 .35rem 0; font-size: 2rem;}
      .muted {color:#666;}
      div[data-testid="stMetric"] {background:white; border:1px solid #eee; padding:12px; border-radius:12px;}
    </style>
    <div class="hero">
      <h1>Tarifname Atölyesi</h1>
      <div class="muted">BBF yükle, unsurları kontrol et, sistem/yöntem istemlerini üret ve Word tarifnamesini indir.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if "extracted" not in st.session_state:
    st.session_state.extracted = None
if "literature" not in st.session_state:
    st.session_state.literature = []
if "draft" not in st.session_state:
    st.session_state.draft = None

with st.sidebar:
    st.header("Proje ayarları")
    output_name = st.text_input("Çıktı dosyası", value="Tarifname_181710.docx")
    claim_mode = st.selectbox(
        "İstem yapısı",
        ["Otomatik analiz et", "Yalnızca sistem", "Sistem ve yöntem"],
    )
    do_research = st.checkbox("Literatür araştırması yapılsın", value=False)
    max_docs = st.slider("Benzer doküman sayısı", 1, 10, 5, disabled=not do_research)
    template_upload = st.file_uploader("İsteğe bağlı DOCX şablonu", type=["docx"])
    st.divider()
    model = os.getenv("OPENAI_MODEL", "gpt-5.6")
    st.caption(f"Model: {model}")
    if os.getenv("OPENAI_API_KEY"):
        st.success("API anahtarı tanımlı")
    else:
        st.warning("OPENAI_API_KEY tanımlı değil")

uploaded = st.file_uploader("BBF dosyasını yükleyin", type=["docx", "doc", "pdf", "txt"])

col1, col2, col3 = st.columns(3)
col1.metric("1", "BBF analizi")
col2.metric("2", "İstem ve metin kontrolü")
col3.metric("3", "Word çıktısı")

if uploaded:
    data = uploaded.getvalue()
    try:
        source_text = extract_text(uploaded.name, data)
        with st.expander("BBF'den çıkarılan ham metin"):
            st.text_area("Metin", source_text, height=280, label_visibility="collapsed")
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    if st.button("BBF'yi analiz et", type="primary", use_container_width=True):
        with st.spinner("Unsurlar ve işlem adımları çıkarılıyor..."):
            st.session_state.extracted = ask_json(extraction_prompt(source_text))
            st.session_state.draft = None
            st.session_state.literature = []

if st.session_state.extracted:
    ex = st.session_state.extracted
    st.subheader("BBF kontrolü")
    c1, c2 = st.columns([2, 1])
    ex["title"] = c1.text_input("Buluş başlığı", value=ex.get("title", ""))
    c2.info("Yöntem temeli: " + ("Var" if ex.get("has_method_basis") else "Yok"))
    st.caption(ex.get("method_basis_reason", ""))

    st.markdown("**Unsurlar**")
    elements_df = pd.DataFrame(ex.get("elements") or [], columns=["number", "name", "function"])
    edited_elements = st.data_editor(elements_df, num_rows="dynamic", use_container_width=True)
    ex["elements"] = edited_elements.fillna("").to_dict("records")

    st.markdown("**Yöntem işlem adımları**")
    steps_df = pd.DataFrame(ex.get("method_steps") or [], columns=["number", "text"])
    edited_steps = st.data_editor(steps_df, num_rows="dynamic", use_container_width=True)
    ex["method_steps"] = edited_steps.fillna("").to_dict("records")

    uncertainties = ex.get("uncertainties") or []
    if uncertainties:
        st.warning("BBF'de netleştirilmesi gereken noktalar:\n\n" + "\n".join(f"- {x}" for x in uncertainties))

    if do_research and st.button("Benzer patent dokümanlarını araştır", use_container_width=True):
        with st.spinner("Patent literatürü araştırılıyor..."):
            result = ask_json(literature_prompt(ex, max_docs), use_web=True)
            st.session_state.literature = result.get("documents") or []

    if st.session_state.literature:
        st.markdown("**Literatür adayları**")
        lit_df = pd.DataFrame(st.session_state.literature)
        lit_df.insert(0, "seç", False)
        selected = st.data_editor(lit_df, use_container_width=True, disabled=[c for c in lit_df.columns if c != "seç"])
        st.session_state.selected_literature = selected[selected["seç"]].drop(columns=["seç"]).to_dict("records")
    else:
        st.session_state.selected_literature = []

    if st.button("Tarifname taslağını oluştur", type="primary", use_container_width=True):
        mode = claim_mode
        if claim_mode == "Otomatik analiz et":
            mode = "Sistem ve yöntem" if ex.get("has_method_basis") else "Yalnızca sistem"
        with st.spinner("Tarifname ve istemler hazırlanıyor..."):
            st.session_state.draft = ask_json(
                drafting_prompt(ex, mode, st.session_state.get("selected_literature", []))
            )

if st.session_state.draft:
    draft = st.session_state.draft
    st.subheader("Nihai kontrol")
    findings = validate_draft(draft)
    for finding in findings:
        if finding["level"] == "Hata":
            st.error(finding["message"])
        elif finding["level"] == "Uyarı":
            st.warning(finding["message"])
        else:
            st.success(finding["message"])

    tabs = st.tabs(["İstemler", "Referanslar", "Detaylı açıklama", "Ham JSON"])
    with tabs[0]:
        sc = draft.get("system_claim") or {}
        st.markdown("**Bağımsız sistem istemi**")
        st.write(sc.get("preamble", ""))
        for x in sc.get("elements") or []:
            st.write("• " + x)
        st.write(sc.get("closing", ""))
        if draft.get("method_claim"):
            st.markdown("**Bağımsız yöntem istemi**")
            mc = draft["method_claim"]
            st.write(mc.get("preamble", ""))
            for x in mc.get("steps") or []:
                st.write("• " + x)
            st.write(mc.get("closing", ""))
    with tabs[1]:
        st.dataframe(pd.DataFrame(draft.get("elements") or []), use_container_width=True)
        st.dataframe(pd.DataFrame(draft.get("method_steps") or []), use_container_width=True)
    with tabs[2]:
        for p in draft.get("detailed_paragraphs") or []:
            st.write(p)
    with tabs[3]:
        st.json(draft)

    template_path = DEFAULT_TEMPLATE
    tmp_template = None
    if template_upload:
        tmp_template = BASE_DIR / ".uploaded_template.docx"
        tmp_template.write_bytes(template_upload.getvalue())
        template_path = tmp_template

    try:
        docx_bytes = build_docx(draft, template_path)
        safe_name = output_name if output_name.lower().endswith(".docx") else output_name + ".docx"
        st.download_button(
            "Word tarifnamesini indir",
            data=docx_bytes,
            file_name=safe_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
    finally:
        if tmp_template and tmp_template.exists():
            tmp_template.unlink()
