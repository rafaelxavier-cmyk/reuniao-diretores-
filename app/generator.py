"""
generator.py – Converte um dicionário de pauta em .pptx com identidade visual Matriz Educação.

Regras tipográficas aplicadas:
  • Títulos principais:     30pt bold
  • Títulos de cards/seção: 16pt bold
  • Corpo / bullets:        13pt
  • Labels / badges:        10pt bold
  • Footer / info secundária: 10pt
  • NUNCA usar fonte < 10pt em conteúdo visível
"""

from __future__ import annotations
import io
import re
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── PALETA ─────────────────────────────────────────────────────────────────
NAVY      = RGBColor(0x0E, 0x3D, 0x52)
TEAL      = RGBColor(0x5D, 0xD4, 0xB4)
DARK_NAVY = RGBColor(0x0A, 0x2D, 0x3D)
LIGHT_BG  = RGBColor(0xF4, 0xF7, 0xF8)
AMBER     = RGBColor(0xD9, 0x77, 0x06)
GREEN     = RGBColor(0x16, 0xA3, 0x4A)
RED_SOFT  = RGBColor(0xDC, 0x26, 0x26)
GRAY_TXT  = RGBColor(0x6B, 0x72, 0x80)
DARK_TXT  = RGBColor(0x1F, 0x2A, 0x37)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
SEPARATOR = RGBColor(0xE5, 0xE7, 0xEB)

ACCENTS = [TEAL, NAVY, TEAL, NAVY, AMBER, TEAL, NAVY, AMBER]

# ── PARÂMETROS DE QUALIDADE (sobrescritos por quality_loop) ─────────────────
_GEN_PARAMS: dict = {}

def _get_param(key: str, default):
    return _GEN_PARAMS.get(key, default)

# ── DIMENSÕES ───────────────────────────────────────────────────────────────
W, H        = 12192000, 6858000
SIDEBAR_W   = 164592
HEADER_H    = 1097280
CONTENT_L   = 548640
CONTENT_W   = 10058400
FULL_W      = W - CONTENT_L - 200000         # conteúdo até margem direita
LOGO_L, LOGO_T = 465055, -116672
LOGO_W, LOGO_H = 2864869, 1053344
CO_L, CO_T  = 8686800, 320040
SEC_T       = 566928
SUP_T       = 1051560
TITLE_T     = 1325880
TITLE_H     = 650000
CONTENT_TOP = 2150000                        # Y onde o conteúdo começa
FOOTER_T    = 6537960
SAFE_BOTTOM = 6450000                        # margem segura antes do footer

# ── TIPOGRAFIA (em pontos) ──────────────────────────────────────────────────
T = {
    "title":        30,
    "card_title":   16,
    "card_body":    13,
    "bullet":       13,
    "label":        10,
    "badge":        10,
    "supertitle":   11,
    "company":      10,
    "section_lbl":   9,
    "footer":       10,
    "kpi_value":    24,
    "kpi_label":    10,
    "table_header": 11,
    "table_body":   11,
    "tag":          10,
}

# ── HELPERS PRIMITIVOS ──────────────────────────────────────────────────────

def _load_logo(assets_dir: Path) -> bytes | None:
    for name in ("logo.png", "logo.jpg", "logo.jpeg"):
        p = assets_dir / name
        if p.exists():
            return p.read_bytes()
    return None


def _rect(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(1, Emu(l), Emu(t), Emu(w), Emu(h))
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def _txt(slide, l, t, w, h, text, sz,
         bold=False, color=DARK_TXT, align=PP_ALIGN.LEFT,
         italic=False, wrap=True):
    tb = slide.shapes.add_textbox(Emu(l), Emu(t), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text       = text
    r.font.size  = Pt(sz)
    r.font.bold  = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name  = "Calibri"
    return tb


def _bullets(slide, l, t, w, h, items,
             sz=None, color=DARK_TXT, spacing=8,
             max_items: int | None = None):
    """Renderiza lista de bullets. Trunca em max_items."""
    if not items:
        return
    sz = sz or T["bullet"]
    display = list(items)
    if max_items and len(display) > max_items:
        display = display[:max_items]
        display.append(f"… e mais {len(items) - max_items} itens")

    tb = slide.shapes.add_textbox(Emu(l), Emu(t), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(display):
        if isinstance(item, dict):
            text = item.get("text", "")
            s    = item.get("sz", sz)
            c    = item.get("color", color)
            bd   = item.get("bold", False)
            bul  = item.get("bullet", True)
        else:
            text = str(item)
            s, c, bd, bul = sz, color, False, True

        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(spacing)
        r = p.add_run()
        r.text       = ("• " if bul else "") + text
        r.font.size  = Pt(s)
        r.font.bold  = bd
        r.font.color.rgb = c
        r.font.name  = "Calibri"


def _add_logo(slide, logo_blob):
    if logo_blob:
        slide.shapes.add_picture(
            io.BytesIO(logo_blob),
            Emu(LOGO_L), Emu(LOGO_T), Emu(LOGO_W), Emu(LOGO_H))


def _header(slide, logo_blob, supertitle: str, title: str,
            section: str = "Reunião de Diretores",
            footer: str = ""):
    """Header padrão presente em todos os slides de conteúdo."""
    # Barra lateral
    _rect(slide, 0, 0, SIDEBAR_W, H, NAVY)
    _rect(slide, 0, 0, SIDEBAR_W, HEADER_H, TEAL)
    # Nome empresa + seção (canto superior direito)
    _txt(slide, CO_L, CO_T, 3200400, 274320,
         "MATRIZ EDUCAÇÃO", T["company"], bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
    _txt(slide, CO_L, SEC_T, 3200400, 228600,
         section, T["section_lbl"], color=GRAY_TXT, align=PP_ALIGN.RIGHT)
    # Separador
    _rect(slide, CONTENT_L, 868680, 11338560, 9144, SEPARATOR)
    # Supertítulo + título principal
    _txt(slide, CONTENT_L, SUP_T, FULL_W, 274320,
         supertitle.upper(), T["supertitle"], bold=True, color=TEAL)
    _txt(slide, CONTENT_L, TITLE_T, FULL_W, TITLE_H,
         title, T["title"], bold=True, color=DARK_NAVY)
    # Footer
    if footer:
        _txt(slide, CONTENT_L, FOOTER_T, 7315200, 274320,
             footer, T["footer"], color=GRAY_TXT)
    _add_logo(slide, logo_blob)


def _tag(slide, l, t, w, h, text, bg):
    _rect(slide, l, t, w, h, bg)
    _txt(slide, l, t, w, h, text, T["tag"], bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)


def _card(slide, l, t, w, h, accent, bg=LIGHT_BG):
    """Fundo de card com barra de acento no topo."""
    _rect(slide, l, t, w, h, bg)
    _rect(slide, l, t, w, 73152, accent)


def _available_content_height():
    return SAFE_BOTTOM - CONTENT_TOP


# ══════════════════════════════════════════════════════════════════════════
# SLIDES
# ══════════════════════════════════════════════════════════════════════════

def _slide_capa(prs, blank, logo_blob, pauta: dict, footer: str):
    s = prs.slides.add_slide(blank)

    # Fundo
    _rect(s, 0, 0, W, H, NAVY)
    _rect(s, SIDEBAR_W, 0, W - SIDEBAR_W, H, WHITE)
    _rect(s, 0, 0, SIDEBAR_W, H, NAVY)
    _rect(s, 0, 0, SIDEBAR_W, HEADER_H, TEAL)

    # Faixas decorativas diagonais (efeito visual)
    _rect(s, W - 1500000, 0, 1500000, H, RGBColor(0xF0, 0xFB, 0xF7))  # verde claro
    _rect(s, W - 400000, 0, 400000, H, TEAL)                            # borda teal
    _rect(s, W - 800000, 0, 9144, H, SEPARATOR)                         # linha

    _add_logo(s, logo_blob)

    _txt(s, CO_L, CO_T, 3000000, 274320,
         "MATRIZ EDUCAÇÃO", T["company"], bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
    _rect(s, CONTENT_L, 868680, W - SIDEBAR_W - 1600000, 9144, SEPARATOR)

    # Bloco central de título
    _rect(s, CONTENT_L, 1500000, W - SIDEBAR_W - 1800000, 1000000, DARK_NAVY)
    _rect(s, CONTENT_L, 1500000, 120000, 1000000, TEAL)
    _txt(s, CONTENT_L + 220000, 1570000, 8000000, 260000,
         "REUNIÃO DE DIRETORES", T["supertitle"], bold=True, color=TEAL)

    data_fmt = pauta.get("data", "") or "—"
    _txt(s, CONTENT_L + 220000, 1820000, 8000000, 620000,
         data_fmt, 32, bold=True, color=WHITE)

    # Objetivo
    if pauta.get("objetivo"):
        _txt(s, CONTENT_L, 2650000, W - SIDEBAR_W - 1800000, 350000,
             pauta["objetivo"], 13, color=GRAY_TXT, italic=True, wrap=True)

    # Linha divisória
    _rect(s, CONTENT_L, 3100000, W - SIDEBAR_W - 1800000, 9144, SEPARATOR)

    # Cards de info
    agenda = pauta.get("agenda", [])
    total_min = sum(
        int(re.search(r"\d+", a["tempo"]).group())
        for a in agenda if a.get("tempo") and re.search(r"\d+", a["tempo"])
    )
    for i, (lbl, val, clr) in enumerate([
        ("DURAÇÃO TOTAL", f"{total_min} min" if total_min else "—", TEAL),
        ("FORMATO", "Presencial · Reunião de Diretores", NAVY),
    ]):
        ox = CONTENT_L + i * 5100000
        _rect(s, ox, 3250000, 4800000, 560000, LIGHT_BG)
        _rect(s, ox, 3250000, 80000, 560000, clr)
        _txt(s, ox + 180000, 3275000, 4400000, 240000, lbl, T["label"], bold=True, color=clr)
        _txt(s, ox + 180000, 3530000, 4400000, 280000, val, 13, color=DARK_TXT)

    # Pauta resumida
    _rect(s, CONTENT_L, 3950000, W - SIDEBAR_W - 1800000, 9144, TEAL)
    _txt(s, CONTENT_L, 4010000, 5000000, 240000, "PAUTA", T["supertitle"], bold=True, color=TEAL)

    real_items = [a for a in agenda if a["num"] > 0]
    half = (len(real_items) + 1) // 2
    for ci, chunk in enumerate([real_items[:half], real_items[half:]]):
        lx = CONTENT_L + ci * 5100000
        for ri, item in enumerate(chunk):
            _txt(s, lx, 4280000 + ri * 390000, 4800000, 380000,
                 f"{item['num']:02d}  {item['nome']}" +
                 (f"  —  {item['tempo']}" if item.get("tempo") else ""),
                 13, bold=True, color=DARK_TXT)

    _txt(s, CONTENT_L, FOOTER_T, 7315200, 274320, footer, T["footer"], color=GRAY_TXT)


def _slide_agenda(prs, blank, logo_blob, pauta: dict, footer: str):
    s = prs.slides.add_slide(blank)
    _header(s, logo_blob,
            f"Reunião de Diretores · {pauta.get('data', '')}",
            "Pauta do Dia", footer=footer)

    agenda = pauta.get("agenda", [])
    if not agenda:
        return

    # ── Layout dinâmico: ajusta ROW_H para caber todos os itens ────────────
    ENC_RESERVE = 320000                               # barra de encerramento + margem
    AVAIL_H     = SAFE_BOTTOM - ENC_RESERVE - CONTENT_TOP - 60000
    ROW_GAP     = 40000
    INT_H       = 220000                               # altura da linha de intervalo

    n_regular  = sum(1 for a in agenda if a["num"] > 0)
    n_interval = sum(1 for a in agenda if a["num"] == 0)

    # Espaço consumido pelos intervalos
    interval_total = n_interval * (INT_H + ROW_GAP)
    avail_regular  = AVAIL_H - interval_total

    # ROW_H ideal: divide igualmente o espaço restante
    ROW_H_DEFAULT = 530000
    ROW_H_MIN     = 220000
    if n_regular > 0:
        ROW_H = max(ROW_H_MIN, min(ROW_H_DEFAULT,
                                   (avail_regular // n_regular) - ROW_GAP))
    else:
        ROW_H = ROW_H_DEFAULT

    # Escala de fonte proporcional ao ROW_H (mín 11pt)
    font_scale = ROW_H / ROW_H_DEFAULT
    nome_sz    = max(11, int(16 * font_scale))
    tag_h      = max(220000, min(310000, int(ROW_H * 0.55)))

    y = CONTENT_TOP

    for item in agenda:
        if item["num"] == 0:
            _rect(s, CONTENT_L, y + ROW_GAP // 2, FULL_W, INT_H,
                  RGBColor(0xEF, 0xF6, 0xF3))
            _rect(s, CONTENT_L, y + ROW_GAP // 2, 80000, INT_H, SEPARATOR)
            _txt(s, CONTENT_L + 170000, y + ROW_GAP // 2 + 20000,
                 5000000, INT_H,
                 f"  INTERVALO  —  {item['tempo']}", T["label"],
                 bold=True, color=GRAY_TXT)
            y += INT_H + ROW_GAP
            continue

        accent  = ACCENTS[(item["num"] - 1) % len(ACCENTS)]
        name_top = y + max(20000, (ROW_H - int(nome_sz * 16000)) // 2)

        _rect(s, CONTENT_L, y, FULL_W, ROW_H, LIGHT_BG)
        _rect(s, CONTENT_L, y, 80000, ROW_H, accent)
        _txt(s, CONTENT_L + 80000, y + 10000, 380000, ROW_H - 20000,
             f"{item['num']:02d}", nome_sz, bold=True, color=accent,
             align=PP_ALIGN.CENTER)
        _txt(s, CONTENT_L + 520000, name_top, 8000000,
             min(300000, ROW_H - 30000),
             item["nome"], nome_sz, bold=True, color=DARK_NAVY)
        if item.get("tempo"):
            _tag(s,
                 CONTENT_L + FULL_W - 1200000,
                 y + (ROW_H - tag_h) // 2,
                 1150000, tag_h,
                 item["tempo"], NAVY)
        y += ROW_H + ROW_GAP

    # Barra de encerramento — sempre abaixo do último item, nunca sobreposta
    enc_y = max(y + 20000, SAFE_BOTTOM - 310000)
    _rect(s, CONTENT_L, enc_y, FULL_W, 300000, TEAL)
    _txt(s, CONTENT_L + 200000, enc_y + 100000, 5000000, 200000,
         "ENCERRAMENTO  ·  FIM DA REUNIÃO", T["label"],
         bold=True, color=DARK_NAVY)


def _slide_section_generic(prs, blank, logo_blob, sec: dict, footer: str, idx: int):
    accent = ACCENTS[idx % len(ACCENTS)]
    supertitle = f"{sec['num']:02d} · {sec['titulo']}"
    if sec.get("tempo"):
        supertitle += f"  —  {sec['tempo']}"

    items_all  = [it for it in sec.get("items", []) if it.strip()]
    subsecoes  = [sub for sub in sec.get("subsecoes", []) if sub.get("items") or sub.get("titulo")]

    if subsecoes:
        _slide_with_subs(prs, blank, logo_blob, sec, supertitle, accent, subsecoes, items_all, footer)
    elif items_all:
        _slide_bullets_only(prs, blank, logo_blob, sec, supertitle, accent, items_all, footer)
    else:
        # Slide vazio — só título
        s = prs.slides.add_slide(blank)
        _header(s, logo_blob, supertitle, sec["titulo"], footer=footer)
        _txt(s, CONTENT_L, CONTENT_TOP + 400000, FULL_W, 300000,
             "Conteúdo a ser apresentado na reunião",
             T["card_body"], color=GRAY_TXT, italic=True)


def _slide_with_subs(prs, blank, logo_blob, sec, supertitle, accent,
                     subsecoes, extra_items, footer):
    s = prs.slides.add_slide(blank)
    _header(s, logo_blob, supertitle, sec["titulo"], footer=footer)

    n    = min(len(subsecoes), 4)   # máx 4 cards por slide
    cols = 2 if n > 2 else n
    rows = (n + 1) // 2

    CARD_GAP = 80000
    BW = (FULL_W - (cols - 1) * CARD_GAP) // cols
    avail_h = SAFE_BOTTOM - CONTENT_TOP
    BH_natural = (avail_h - (rows - 1) * CARD_GAP) // rows
    BH = min(BH_natural, 2600000)  # allow taller cards, cap only to avoid extreme cases
    BH = max(BH, 1400000)

    for i, sub in enumerate(subsecoes[:4]):
        row, col = divmod(i, cols)
        bx = CONTENT_L + col * (BW + CARD_GAP)
        by = CONTENT_TOP + row * (BH + CARD_GAP)

        _card(s, bx, by, BW, BH, accent)
        _rect(s, bx, by + 73152, 80000, BH - 73152, LIGHT_BG)  # margem esq

        # Título da sub-seção
        title_text = sub["titulo"]
        # Allow 2 lines for long titles
        title_h = 450000 if len(title_text) > 35 else 320000
        _txt(s, bx + 170000, by + 130000, BW - 270000, title_h,
             title_text, T["card_title"], bold=True, color=DARK_NAVY, wrap=True)

        # Bullets
        items_disp = sub.get("items", [])
        max_items  = _get_param('max_bullets_per_card', max(3, int((BH - 600000) // (T["bullet"] * 20000))))
        spacing_factor = _get_param('extra_spacing_factor', 1.0)
        bullets_top = by + 130000 + title_h + 60000
        bullets_h   = by + BH - 80000 - bullets_top  # fill card leaving 80000 bottom padding
        _bullets(s, bx + 170000, bullets_top,
                 BW - 280000, max(bullets_h, 500000),
                 items_disp, sz=T["bullet"], spacing=int(8 * spacing_factor), max_items=max_items)

    # Extra items (overflow de sub-seções > 4)
    if len(subsecoes) > 4:
        _slide_with_subs(prs, blank, logo_blob, sec, supertitle, accent,
                         subsecoes[4:], extra_items, footer)


def _slide_bullets_only(prs, blank, logo_blob, sec, supertitle, accent,
                        items, footer):
    # Se muitos itens, divide em 2 colunas ou páginas
    MAX_PER_SLIDE = 8
    for chunk_start in range(0, len(items), MAX_PER_SLIDE):
        chunk = items[chunk_start:chunk_start + MAX_PER_SLIDE]
        s = prs.slides.add_slide(blank)
        part_label = supertitle + (f" (cont.)" if chunk_start > 0 else "")
        _header(s, logo_blob, part_label, sec["titulo"], footer=footer)

        CT = CONTENT_TOP

        # Bloco de destaque (subtítulo)
        if sec.get("subtitulo") and sec["subtitulo"] != sec["titulo"] and chunk_start == 0:
            obj_txt  = sec.get("objetivo_label", "") or sec.get("subtitulo", "")
            block_h  = 850000 if (obj_txt and len(obj_txt) > 60) else 730000
            _rect(s, CONTENT_L, CT, FULL_W, block_h, DARK_NAVY)
            _rect(s, CONTENT_L, CT, 110000, block_h, accent)
            _txt(s, CONTENT_L + 210000, CT + 120000, FULL_W - 300000, 220000,
                 "OBJETIVO", T["label"], bold=True, color=accent)
            _txt(s, CONTENT_L + 210000, CT + 360000, FULL_W - 300000, block_h - 370000,
                 sec["subtitulo"], T["card_title"], bold=True, color=WHITE, wrap=True)
            CT += block_h + 100000

        avail = SAFE_BOTTOM - CT - 100000

        if len(chunk) > 4:
            half = (len(chunk) + 1) // 2
            col_w = (FULL_W - 150000) // 2
            for ci, col_items in enumerate([chunk[:half], chunk[half:]]):
                lx = CONTENT_L + ci * (col_w + 150000)
                _bullets(s, lx, CT + 150000, col_w, avail - 150000,
                         col_items, sz=T["bullet"], spacing=10)
        else:
            _bullets(s, CONTENT_L, CT + 150000, FULL_W, avail - 150000,
                     chunk, sz=T["bullet"], spacing=14)


def _slide_encerramento(prs, blank, logo_blob, pauta: dict, footer: str):
    s = prs.slides.add_slide(blank)
    _rect(s, 0, 0, W, H, NAVY)
    _rect(s, 0, 0, SIDEBAR_W, H, DARK_NAVY)
    _rect(s, 0, 0, SIDEBAR_W, HEADER_H, TEAL)
    _rect(s, W - 400000, 0, 400000, H, TEAL)
    _add_logo(s, logo_blob)

    _txt(s, CO_L, CO_T, 3200400, 274320,
         "MATRIZ EDUCAÇÃO", T["company"], bold=True, color=TEAL, align=PP_ALIGN.RIGHT)

    _rect(s, CONTENT_L, 1900000, FULL_W, 1000000, RGBColor(0x16, 0x4E, 0x63))
    _rect(s, CONTENT_L, 1900000, 120000, 1000000, TEAL)
    _txt(s, CONTENT_L + 220000, 1970000, FULL_W - 300000, 260000,
         "ENCERRAMENTO", T["supertitle"], bold=True, color=TEAL)
    _txt(s, CONTENT_L + 220000, 2220000, FULL_W - 300000, 580000,
         f"Reunião de Diretores  —  {pauta.get('data', '')}",
         24, bold=True, color=WHITE)

    enc = [
        ("Dúvidas",                "Espaço aberto para perguntas e esclarecimentos"),
        ("Temas Diversos",         "Assuntos não contemplados na pauta"),
        ("Reforço dos Combinados", "Revisão dos acordos e responsabilidades da reunião"),
        ("Prioridades da Semana",  "Alinhamento do que cada área executará nos próximos dias"),
    ]
    EW  = (FULL_W - 80000) // 2
    EH  = 780000
    base = RGBColor(0x16, 0x4E, 0x63)
    for i, (title, desc) in enumerate(enc):
        row, col = divmod(i, 2)
        ex = CONTENT_L + col * (EW + 80000)
        ey = 3100000 + row * (EH + 73152)
        _rect(s, ex, ey, EW, EH, base)
        _rect(s, ex, ey, EW, 73152, TEAL)
        _rect(s, ex, ey + 73152, 80000, EH - 73152, DARK_NAVY)
        _txt(s, ex + 170000, ey + 140000, EW - 250000, 300000,
             title, T["card_title"], bold=True, color=TEAL, wrap=True)
        _txt(s, ex + 170000, ey + 460000, EW - 250000, 300000,
             desc, T["card_body"], color=RGBColor(0xB8, 0xEB, 0xDA), wrap=True)

    _txt(s, CONTENT_L, FOOTER_T, 7315200, 274320, footer, T["footer"], color=GRAY_TXT)
    _txt(s, 8686800, FOOTER_T, 3200400, 274320,
         "FIM", T["footer"], bold=True, color=TEAL, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════

def generate(pauta: dict, output_path: str,
             assets_dir: str | Path | None = None) -> str:
    global _GEN_PARAMS
    _GEN_PARAMS = pauta.get('_gen_params', {})

    if assets_dir is None:
        assets_dir = Path(__file__).parent / "assets"
    assets_dir = Path(assets_dir)
    logo_blob  = _load_logo(assets_dir)

    prs = Presentation()
    prs.slide_width  = Emu(W)
    prs.slide_height = Emu(H)
    blank = prs.slide_layouts[6]

    data   = pauta.get("data", "")
    footer = f"Matriz Educação  ·  Reunião de Diretores{' — ' + data if data else ''}"

    _slide_capa(prs, blank, logo_blob, pauta, footer)

    if pauta.get("agenda"):
        _slide_agenda(prs, blank, logo_blob, pauta, footer)

    for i, sec in enumerate(pauta.get("secoes", [])):
        _slide_section_generic(prs, blank, logo_blob, sec, footer, i)

    _slide_encerramento(prs, blank, logo_blob, pauta, footer)

    prs.save(output_path)
    return output_path
