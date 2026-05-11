"""
app.py – Gerador de apresentações Matriz Educação

Modos:
  • Upload → sobe DOCX/PDF e gera o PPT automaticamente
  • Criar   → assistente guiado que constrói a pauta passo a passo
"""

import io
import os
import re
import tempfile
from copy import deepcopy
from datetime import datetime
from pathlib import Path

import streamlit as st

BASE_DIR   = Path(__file__).parent.parent
ASSETS_DIR = Path(__file__).parent / "assets"
SAIDA_DIR  = BASE_DIR / "saida"
SAIDA_DIR.mkdir(exist_ok=True)


# ══════════════════════════════════════════════════════════════════════════
# FUNÇÕES AUXILIARES
# ══════════════════════════════════════════════════════════════════════════

def _build_filename(pauta: dict) -> str:
    data = pauta.get("data", "").strip()
    data_clean = re.sub(r"[^\w\d]", "-", data) if data else datetime.now().strftime("%d-%m-%Y")
    return f"Reunião de Diretores - {data_clean}.pptx"


def _render_preview(pauta: dict):
    st.markdown("### Estrutura detectada")
    mc1, mc2 = st.columns(2)
    with mc1:
        st.metric("Data", pauta.get("data") or "—")
    with mc2:
        n_sec   = len(pauta.get("secoes", []))
        st.metric("Slides a gerar", str(2 + n_sec + 1))

    if pauta.get("objetivo"):
        st.info(f"**Objetivo:** {pauta['objetivo']}")

    if pauta.get("agenda"):
        st.markdown("**Agenda:**")
        for item in pauta["agenda"]:
            if item["num"] == 0:
                st.markdown(
                    f"<div style='padding:4px 16px;color:#9CA3AF;font-size:12px;'>"
                    f"└── Intervalo – {item['tempo']}</div>",
                    unsafe_allow_html=True)
            else:
                accent = "#5DD4B4" if item["num"] % 2 == 1 else "#0E3D52"
                st.markdown(
                    f"<div style='border-left:4px solid {accent};padding:6px 12px;"
                    f"background:#fff;border-radius:0 6px 6px 0;margin-bottom:6px;'>"
                    f"<b style='color:{accent}'>{item['num']:02d}</b> "
                    f"<span style='color:#0A2D3D;font-weight:600'>{item['nome']}</span> "
                    f"<span style='color:#6B7280;font-size:12px'>· {item.get('tempo','')}</span>"
                    f"</div>",
                    unsafe_allow_html=True)

    if pauta.get("secoes"):
        st.markdown("**Conteúdo por seção:**")
        for sec in pauta["secoes"]:
            label = (f"**{sec['num']:02d}. {sec['titulo']}**" +
                     (f"  ·  _{sec['tempo']}_" if sec.get("tempo") else ""))
            with st.expander(label):
                for it in (sec.get("items") or [])[:8]:
                    st.markdown(f"• {it}")
                for sub in sec.get("subsecoes") or []:
                    if sub.get("titulo"):
                        st.markdown(f"**{sub['titulo']}**")
                    for it in (sub.get("items") or [])[:4]:
                        st.markdown(f"  – {it}")


def _run_generation(pauta: dict):
    """Gera PPT com loop de qualidade e retorna (bytes, filename, audit_report)."""
    from quality_loop import run_quality_loop

    out_name = _build_filename(pauta)
    out_path = str(SAIDA_DIR / out_name)

    status_box = st.empty()
    log_box    = st.empty()

    def _progress(step: str, iteration: int, score):
        if score is None:
            status_box.info(f"Iteração {iteration}/{4} — Gerando slides...")
        else:
            status_box.info(f"Iteração {iteration}/{4} concluída — Score: {score}/100")

    pptx_bytes, audit_report, iterations_log = run_quality_loop(
        pauta, out_path, ASSETS_DIR, progress_callback=_progress
    )

    status_box.empty()

    if len(iterations_log) > 1:
        trail = " → ".join(f"it{l['iteration']}: {l['score']}" for l in iterations_log)
        log_box.caption(f"Loop de qualidade: {trail} (score final: {audit_report.score}/100)")

    return pptx_bytes, out_name, audit_report


def _show_download(pptx_bytes, out_name, report):
    score = report.score
    if score >= 90:   grade = "A"
    elif score >= 75: grade = "B"
    elif score >= 60: grade = "C"
    else:             grade = "D"

    grade_color = {"A": "#16A34A", "B": "#65A30D", "C": "#D97706", "D": "#DC2626"}
    color = grade_color[grade]

    st.markdown(
        f"<div style='background:#ECFDF5;border:1px solid #6EE7B7;border-radius:10px;"
        f"padding:16px 20px;text-align:center;'>"
        f"<div style='font-size:22px;font-weight:700;color:#065F46;'>Apresentacao pronta!</div>"
        f"<div style='margin-top:6px;color:#047857;'>"
        f"Qualidade visual: <b style='color:{color};font-size:18px;'>{grade}</b> "
        f"({score}/100)"
        f"</div></div>",
        unsafe_allow_html=True)

    if report.issues:
        with st.expander(f"{len(report.issues)} observacoes do auditor de layout"):
            for i in report.issues:
                icon = {"error": "🔴", "warning": "🟡"}.get(i.severity, "•")
                st.markdown(f"{icon} **Slide {i.slide}** [{i.kind}]: {i.details}")

    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        label=f"Baixar {out_name}",
        data=pptx_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )


# ══════════════════════════════════════════════════════════════════════════
# MODO: CRIAR PAUTA GUIADA
# ══════════════════════════════════════════════════════════════════════════

def _wizard():
    """Assistente passo a passo para construir uma pauta do zero."""
    st.markdown("## ✏️ Criar Pauta — Assistente Guiado")
    st.caption("Preencha cada etapa para montar a melhor pauta possível. "
               "Campos marcados com * são obrigatórios.")

    # Estado do wizard
    if "wiz" not in st.session_state:
        st.session_state.wiz = {
            "step": 1,
            "data": "", "objetivo": "", "formato": "Presencial",
            "topics": [],          # lista de {num, nome, tempo, responsavel}
            "secoes": [],          # conteúdo detalhado de cada tópico
        }
    wiz = st.session_state.wiz

    # ── PROGRESS BAR ───────────────────────────────────────────────────────
    steps = ["Identificação", "Tópicos", "Conteúdo", "Revisão"]
    prog_cols = st.columns(len(steps))
    for i, (col, label) in enumerate(zip(prog_cols, steps)):
        active   = (i + 1 == wiz["step"])
        done     = (i + 1 < wiz["step"])
        bg = "#5DD4B4" if active else ("#0E3D52" if done else "#E5E7EB")
        txt_c = "#0A2D3D" if active else ("#FFFFFF" if done else "#9CA3AF")
        col.markdown(
            f"<div style='text-align:center;background:{bg};border-radius:8px;"
            f"padding:8px;color:{txt_c};font-weight:{'700' if active else '400'};'>"
            f"{'✓ ' if done else ''}{label}</div>",
            unsafe_allow_html=True)

    st.markdown("---")

    # ── ETAPA 1: IDENTIFICAÇÃO ─────────────────────────────────────────────
    if wiz["step"] == 1:
        st.markdown("### Etapa 1 — Identificação da Reunião")

        c1, c2 = st.columns(2)
        with c1:
            wiz["data"] = st.text_input(
                "Data da reunião *",
                value=wiz["data"],
                placeholder="Ex: 07/05/2026",
            )
        with c2:
            wiz["formato"] = st.selectbox(
                "Formato",
                ["Presencial", "Online", "Híbrido"],
                index=["Presencial", "Online", "Híbrido"].index(wiz.get("formato", "Presencial"))
            )

        wiz["objetivo"] = st.text_area(
            "Objetivo da reunião *",
            value=wiz["objetivo"],
            height=90,
            placeholder="Ex: Alinhamento estratégico de ciclo, revisão de metas e fortalecimento da gestão",
        )

        if st.button("Próxima etapa →", type="primary", use_container_width=True):
            if not wiz["data"] or not wiz["objetivo"]:
                st.error("Preencha a data e o objetivo antes de continuar.")
            else:
                wiz["step"] = 2
                st.rerun()

    # ── ETAPA 2: TÓPICOS ───────────────────────────────────────────────────
    elif wiz["step"] == 2:
        st.markdown("### Etapa 2 — Tópicos da Pauta")
        st.caption("Liste os tópicos em ordem, com o tempo e responsável de cada um.")

        # Garante ao menos 1 tópico inicial
        if not wiz["topics"]:
            wiz["topics"] = [_empty_topic(1)]

        # Renderiza cada tópico
        to_remove = None
        for i, topic in enumerate(wiz["topics"]):
            with st.container(border=True):
                c1, c2, c3, c4 = st.columns([0.4, 2.2, 1.2, 1.2])
                with c1:
                    st.markdown(
                        f"<div style='margin-top:28px;font-size:20px;font-weight:700;"
                        f"color:#5DD4B4;text-align:center;'>{topic['num']:02d}</div>",
                        unsafe_allow_html=True)
                with c2:
                    topic["nome"] = st.text_input(
                        "Tópico *", value=topic["nome"],
                        key=f"t_nome_{i}",
                        placeholder="Ex: Operacional, Comercial, Pedagógico…",
                        label_visibility="collapsed")
                with c3:
                    topic["tempo"] = st.text_input(
                        "Tempo", value=topic["tempo"],
                        key=f"t_tempo_{i}",
                        placeholder="Ex: 30 min",
                        label_visibility="collapsed")
                with c4:
                    topic["responsavel"] = st.text_input(
                        "Responsável", value=topic.get("responsavel", ""),
                        key=f"t_resp_{i}",
                        placeholder="Nome",
                        label_visibility="collapsed")
                if len(wiz["topics"]) > 1:
                    if st.button("✕ Remover", key=f"t_rm_{i}"):
                        to_remove = i

        if to_remove is not None:
            wiz["topics"].pop(to_remove)
            _renumber_topics(wiz["topics"])
            st.rerun()

        # Intervalo
        with st.container(border=True):
            col_a, col_b = st.columns([3, 1])
            with col_a:
                wiz["intervalo"] = st.checkbox(
                    "Incluir intervalo",
                    value=wiz.get("intervalo", True))
            with col_b:
                wiz["intervalo_min"] = st.text_input(
                    "Duração", value=wiz.get("intervalo_min", "20 min"),
                    key="intv_dur", label_visibility="collapsed",
                    disabled=not wiz.get("intervalo", True))

        col_add, col_back, col_next = st.columns([2, 1, 1])
        with col_add:
            if st.button("+ Adicionar tópico", use_container_width=True):
                wiz["topics"].append(_empty_topic(len(wiz["topics"]) + 1))
                st.rerun()
        with col_back:
            if st.button("← Voltar", use_container_width=True):
                wiz["step"] = 1; st.rerun()
        with col_next:
            if st.button("Próxima →", type="primary", use_container_width=True):
                valid = [t for t in wiz["topics"] if t["nome"].strip()]
                if not valid:
                    st.error("Adicione ao menos um tópico com nome.")
                else:
                    wiz["topics"] = valid
                    _sync_secoes(wiz)
                    wiz["step"] = 3
                    st.rerun()

    # ── ETAPA 3: CONTEÚDO POR TÓPICO ──────────────────────────────────────
    elif wiz["step"] == 3:
        st.markdown("### Etapa 3 — Conteúdo de cada Tópico")
        st.caption("Para cada tópico, adicione um subtítulo, os pontos principais e "
                   "o objetivo do bloco. Esses dados viram o conteúdo dos slides.")

        for i, sec in enumerate(wiz["secoes"]):
            acc = ["#5DD4B4", "#0E3D52", "#5DD4B4", "#D97706"][i % 4]
            with st.expander(
                f"**{sec['num']:02d}. {sec['titulo']}**" +
                (f" — {sec['tempo']}" if sec.get('tempo') else "") +
                (f" · Resp: {sec.get('responsavel','')}" if sec.get('responsavel') else ""),
                expanded=(i == 0)
            ):
                sec["subtitulo"] = st.text_input(
                    "Subtítulo / Tema central",
                    value=sec.get("subtitulo", ""),
                    key=f"s_sub_{i}",
                    placeholder="Ex: Regras de Ouro – Cultura de Reunião",
                )
                sec["objetivo_label"] = st.text_input(
                    "Objetivo do bloco",
                    value=sec.get("objetivo_label", ""),
                    key=f"s_obj_{i}",
                    placeholder="Ex: Garantir reuniões mais produtivas",
                )
                st.markdown("**Pontos a tratar** (um por linha):")
                items_raw = st.text_area(
                    "Pontos",
                    value="\n".join(sec.get("items", [])),
                    height=130,
                    key=f"s_items_{i}",
                    label_visibility="collapsed",
                    placeholder="• Revisão dos combinados\n• Feedbacks do ciclo anterior\n• Próximos passos",
                )
                sec["items"] = [l.strip().lstrip("•●-▸ ") for l in items_raw.splitlines() if l.strip()]

        col_back, col_next = st.columns(2)
        with col_back:
            if st.button("← Voltar", use_container_width=True):
                wiz["step"] = 2; st.rerun()
        with col_next:
            if st.button("Revisar →", type="primary", use_container_width=True):
                wiz["step"] = 4; st.rerun()

    # ── ETAPA 4: REVISÃO E GERAÇÃO ─────────────────────────────────────────
    elif wiz["step"] == 4:
        st.markdown("### Etapa 4 — Revisão Final")

        pauta = _wiz_to_pauta(wiz)
        _render_preview(pauta)

        col_back, col_gen = st.columns(2)
        with col_back:
            if st.button("← Editar", use_container_width=True):
                wiz["step"] = 3; st.rerun()
        with col_gen:
            if st.button("🚀 Gerar Apresentação", type="primary", use_container_width=True):
                try:
                    pptx_bytes, out_name, report = _run_generation(pauta)
                    st.session_state["wiz_result"] = (pptx_bytes, out_name, report)
                except Exception as e:
                    st.error(f"Erro: {e}")
                    import traceback; st.code(traceback.format_exc())

        if "wiz_result" in st.session_state:
            st.markdown("---")
            _show_download(*st.session_state["wiz_result"])


# ── Helpers do wizard ──────────────────────────────────────────────────────

def _empty_topic(num: int) -> dict:
    return {"num": num, "nome": "", "tempo": "", "responsavel": ""}


def _renumber_topics(topics: list):
    for i, t in enumerate(topics):
        t["num"] = i + 1


def _sync_secoes(wiz: dict):
    """Atualiza wiz['secoes'] preservando conteúdo já preenchido."""
    existing = {s["num"]: s for s in wiz.get("secoes", [])}
    new_secoes = []
    for t in wiz["topics"]:
        n = t["num"]
        if n in existing:
            sec = existing[n]
            sec["titulo"]      = t["nome"]
            sec["tempo"]       = t["tempo"]
            sec["responsavel"] = t.get("responsavel", "")
        else:
            sec = {
                "num":           n,
                "titulo":        t["nome"],
                "tempo":         t["tempo"],
                "responsavel":   t.get("responsavel", ""),
                "subtitulo":     "",
                "objetivo_label":"",
                "items":         [],
                "subsecoes":     [],
            }
        new_secoes.append(sec)
    wiz["secoes"] = new_secoes


def _wiz_to_pauta(wiz: dict) -> dict:
    """Converte o estado do wizard para o formato do gerador."""
    agenda = []
    for t in wiz["topics"]:
        agenda.append({"num": t["num"], "nome": t["nome"], "tempo": t["tempo"]})
        # Intervalo após a metade dos tópicos (se configurado)
    if wiz.get("intervalo"):
        mid = max(1, len(wiz["topics"]) // 2)
        agenda.insert(mid, {"num": 0, "nome": "Intervalo",
                             "tempo": wiz.get("intervalo_min", "20 min")})

    return {
        "titulo":   "REUNIÃO DE DIRETORES",
        "data":     wiz["data"],
        "objetivo": wiz["objetivo"],
        "agenda":   agenda,
        "secoes":   deepcopy(wiz["secoes"]),
    }


# ══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ══════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Pauta → PPT · Matriz Educação",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
  html, body, [data-testid="stAppViewContainer"] { background-color: #F4F7F8; }
  div[data-testid="stDownloadButton"] > button {
      background: #0E3D52 !important; color: white !important;
      border: none !important; border-radius: 8px !important;
      font-weight: 600 !important; font-size: 15px !important; width: 100%;
  }
</style>
""", unsafe_allow_html=True)

# ── HERO ───────────────────────────────────────────────────────────────────
st.markdown("""
<div style="background:linear-gradient(135deg,#0E3D52 60%,#5DD4B4 100%);
     border-radius:12px;padding:28px 40px 20px;margin-bottom:24px;">
  <h1 style="color:#5DD4B4;font-size:26px;margin:0 0 4px;">
    📊  Pauta → Apresentação
  </h1>
  <p style="color:#B8EBDA;font-size:13px;margin:0;">
    Matriz Educação · Gerador automático de slides para reuniões de diretores
  </p>
</div>
""", unsafe_allow_html=True)

# ── TABS ────────────────────────────────────────────────────────────────────
tab_upload, tab_criar = st.tabs(["📂  Upload de Pauta", "✏️  Criar Pauta Guiada"])

# ══════════════════════════════════════════════════════════════════════════
# TAB 1 – UPLOAD
# ══════════════════════════════════════════════════════════════════════════
with tab_upload:
    col_upload, col_preview = st.columns([1, 1.6], gap="large")

    with col_upload:
        with st.container(border=True):
            st.markdown("#### ① Arquivo da pauta")
            uploaded = st.file_uploader(
                "Arraste ou clique para enviar",
                type=["docx", "pdf"],
                label_visibility="collapsed",
            )

        pauta = None

        if uploaded:
            suffix = Path(uploaded.name).suffix
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded.getbuffer())
                tmp_path = tmp.name
            try:
                from parser import parse
                with st.spinner("Lendo documento..."):
                    pauta = parse(tmp_path)
            except Exception as e:
                st.error(f"Erro ao ler o arquivo: {e}")
                st.stop()
            finally:
                try: os.unlink(tmp_path)
                except Exception: pass

            with st.container(border=True):
                st.markdown("#### ② Revise os dados")
                pauta["data"]     = st.text_input("Data", value=pauta.get("data", ""))
                pauta["objetivo"] = st.text_area("Objetivo", value=pauta.get("objetivo", ""),
                                                 height=75)

            with st.container(border=True):
                st.markdown("#### ③ Gere a apresentação")
                if st.button("✨  Gerar", use_container_width=True, type="primary"):
                    try:
                        pptx_bytes, out_name, report = _run_generation(pauta)
                        st.session_state["up_result"] = (pptx_bytes, out_name, report)
                    except Exception as e:
                        st.error(f"Erro: {e}")
                        import traceback; st.code(traceback.format_exc())

            if "up_result" in st.session_state:
                _show_download(*st.session_state["up_result"])

    with col_preview:
        if pauta:
            _render_preview(pauta)
        else:
            st.markdown("""
            <div style="display:flex;flex-direction:column;align-items:center;
                        justify-content:center;height:380px;color:#9CA3AF;text-align:center;">
              <div style="font-size:56px;margin-bottom:12px;">📄</div>
              <div style="font-size:15px;font-weight:600;color:#6B7280;">
                Suba uma pauta para visualizar a estrutura
              </div>
              <div style="font-size:12px;margin-top:6px;">Formatos aceitos: .docx e .pdf</div>
            </div>
            """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# TAB 2 – ASSISTENTE GUIADO
# ══════════════════════════════════════════════════════════════════════════
with tab_criar:
    _wizard()
