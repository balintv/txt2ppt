import io
import re
from pathlib import Path
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_ANCHOR

st.set_page_config(page_title="TXT → PPT", page_icon="🖼️", layout="centered")

st.title("TXT → PPTX generator")
st.caption("Sorokból/bekezdésekből diák. Allítható betűtípus, méret, margók, színek.")

# ---------- Helpers ----------
def read_text_multi_enc(data: bytes) -> str:
    for enc in ("utf-8", "cp1250", "iso-8859-2", "latin2"):
        try:
            return data.decode(enc)
        except Exception:
            continue
    return data.decode("utf-8", errors="replace")

def parse_text(content: str, mode: str, preserve_blanks: bool):
    if mode == "srt":
        blocks = re.split(r"\n\s*\n", content.strip(), flags=re.MULTILINE)
        out = []
        ts_re = re.compile(r"^\s*\d{2}:\d{2}:\d{2},\d{3}\s*-->\s*\d{2}:\d{2}:\d{2},\d{3}\s*$")
        for b in blocks:
            lines = []
            for ln in b.splitlines():
                s = ln.strip()
                if not s:
                    continue
                if s.isdigit():
                    continue
                if ts_re.match(s):
                    continue
                lines.append(s)
            if lines:
                out.append(" ".join(lines))
        return out
    if mode == "para":
        return [b.strip() for b in re.split(r"\n\s*\n", content) if b.strip()]
    if preserve_blanks:
        return content.splitlines()
    return [l.strip() for l in content.splitlines() if l.strip()]

def add_black_slide(prs: Presentation, bg_rgb=(0,0,0)):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*bg_rgb)
    return slide

def add_text_slide(
    prs: Presentation,
    text: str,
    *,
    shrink_to_fit: bool = True,
    m_top_cm=0.13, m_bottom_cm=0.13, m_left_cm=0.25, m_right_cm=0.25,
    font_name="Arial", font_size_pt=44,
    para_left_indent_cm=0.0, para_first_line_indent_cm=0.0,
    align_center=True,
    bg_rgb=(0,0,0),
    font_rgb=(255,255,255)
):
    slide = add_black_slide(prs, bg_rgb=bg_rgb)
    sw, sh = prs.slide_width, prs.slide_height

    # külső margók
    left = Cm(m_left_cm); top = Cm(m_top_cm)
    width = sw - (Cm(m_left_cm) + Cm(m_right_cm))
    height = sh - (Cm(m_top_cm) + Cm(m_bottom_cm))

    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_top = 0; tf.margin_bottom = 0; tf.margin_left = 0; tf.margin_right = 0
    tf.vertical_anchor = MSO_ANCHOR.BOTTOM

    if shrink_to_fit:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    p = tf.paragraphs[0]
    p.text = text
    if para_left_indent_cm:
        p.paragraph_format.left_indent = Cm(para_left_indent_cm)
    if para_first_line_indent_cm:
        p.paragraph_format.first_line_indent = Cm(para_first_line_indent_cm)

    p.font.name = font_name.strip() or "Arial"
    p.font.size = Pt(float(font_size_pt))
    p.font.color.rgb = RGBColor(*font_rgb)

    p.alignment = PP_ALIGN.CENTER if align_center else PP_ALIGN.LEFT

def build_ppt(
    content_items,
    *,
    widescreen=True,
    shrink_to_fit=True,
    blank_slide_on_empty=False,
    m_top_cm=0.13, m_bottom_cm=0.13, m_left_cm=0.25, m_right_cm=0.25,
    font_name="Arial", font_size_pt=44,
    para_left_indent_cm=0.0, para_first_line_indent_cm=0.0,
    align_center=True,
    bg_rgb=(0,0,0),
    font_rgb=(255,255,255)
) -> bytes:
    prs = Presentation()
    if widescreen:
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)

    for item in content_items:
        if blank_slide_on_empty and (item.strip() == ""):
            add_black_slide(prs, bg_rgb=bg_rgb)
        else:
            add_text_slide(
                prs, item,
                shrink_to_fit=shrink_to_fit,
                m_top_cm=m_top_cm, m_bottom_cm=m_bottom_cm,
                m_left_cm=m_left_cm, m_right_cm=m_right_cm,
                font_name=font_name, font_size_pt=font_size_pt,
                para_left_indent_cm=para_left_indent_cm,
                para_first_line_indent_cm=para_first_line_indent_cm,
                align_center=align_center,
                bg_rgb=bg_rgb,
                font_rgb=font_rgb
            )
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

# ---------- UI ----------
st.subheader("Forrás")
tab1, tab2 = st.tabs(["📄 Fájl feltöltés", "📂 Meglévő útvonal"])

uploaded = None
with tab1:
    f = st.file_uploader("Válassz .txt fájlt", type=["txt", "srt"])
    if f is not None:
        data = f.read()
        text = read_text_multi_enc(data)
        uploaded = ("uploaded", text)

with tab2:
    p = st.text_input("Teljes elérési út")
    if p:
        try:
            data = Path(p).read_bytes()
            text = read_text_multi_enc(data)
            uploaded = ("path", text)
        except Exception as e:
            st.error(f"Nem tudtam beolvasni: {e}")

st.subheader("Beállítások")

colA, colB = st.columns(2)

with colA:
    mode = st.selectbox("Felosztás módja", ["line", "para", "srt"], index=0)
    widescreen = st.checkbox("Widescreen 16:9", value=True)
    blank_on_empty = st.checkbox("Üres sor → üres dia (line mód)", value=True)
    shrink = st.checkbox("Hosszú sorok tördelése", value=True)
    align_center = st.checkbox("Vízszintes középre igazítás", value=True)

    # Színek
    bg_hex = st.color_picker("Háttér szín", "#000000")
    font_hex = st.color_picker("Szöveg szín", "#FFFFFF")

with colB:
    font_name = st.text_input("Betűtípus neve", value="Arial")
    font_size_pt = st.number_input("Betűméret (pt)", min_value=8.0, max_value=200.0, value=44.0, step=1.0)

    st.markdown("**Külső margók (cm)**")
    m_top_cm = st.number_input("Felső margó", 0.0, 10.0, 1.0, 0.01)
    m_bottom_cm = st.number_input("Alsó margó", 0.0, 10.0, 1.0, 0.01)
    m_left_cm = st.number_input("Bal margó", 0.0, 20.0, 3.0, 0.01)
    m_right_cm = st.number_input("Jobb margó", 0.0, 20.0, 3.0, 0.01)

    with st.expander("Bekezdés-behúzások"):
        para_left_indent_cm = st.number_input("Bal bekezdés-behúzás (cm)", 0.0, 20.0, 0.0, 0.1)
        para_first_line_indent_cm = st.number_input("Első sor behúzás (cm)", -5.0, 20.0, 0.0, 0.1)

if st.button("PPTX generálása", type="primary", use_container_width=True):
    if not uploaded:
        st.warning("Adj meg forrást (fájl vagy útvonal).")
    else:
        _, raw_text = uploaded
        items = parse_text(raw_text, mode=mode, preserve_blanks=(mode=="line" and blank_on_empty))
        # hex → RGB tuple
        bg_rgb = tuple(int(bg_hex.lstrip("#")[i:i+2], 16) for i in (0,2,4))
        font_rgb = tuple(int(font_hex.lstrip("#")[i:i+2], 16) for i in (0,2,4))
        pptx_bytes = build_ppt(
            items,
            widescreen=widescreen,
            shrink_to_fit=shrink,
            blank_slide_on_empty=blank_on_empty,
            m_top_cm=m_top_cm, m_bottom_cm=m_bottom_cm,
            m_left_cm=m_left_cm, m_right_cm=m_right_cm,
            font_name=font_name, font_size_pt=font_size_pt,
            para_left_indent_cm=para_left_indent_cm,
            para_first_line_indent_cm=para_first_line_indent_cm,
            align_center=align_center,
            bg_rgb=bg_rgb,
            font_rgb=font_rgb
        )
        st.success(f"Siker! {len(items)} dia készült.")
        st.download_button(
            "PPTX letöltése",
            data=pptx_bytes,
            file_name="slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
