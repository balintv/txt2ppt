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
st.caption("Sorokból/bekezdésekből automatikus diák – fekete háttér, Arial 44pt, alulra igazítva.")

# ---------- Helpers ----------
def read_text_multi_enc(data: bytes) -> str:
    # Try common encodings for HU/CEE
    for enc in ("utf-8", "cp1250", "iso-8859-2", "latin2"):
        try:
            return data.decode(enc)
        except Exception:
            continue
    # last resort
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
    # line mode
    if preserve_blanks:
        return content.splitlines()
    return [l.strip() for l in content.splitlines() if l.strip()]

def add_black_slide(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)
    return slide

def add_text_slide(
    prs: Presentation,
    text: str,
    shrink_to_fit: bool = True,
    m_top_cm=1, m_bottom_cm=1, m_left_cm=3, m_right_cm=3
):
    slide = add_black_slide(prs)
    sw, sh = prs.slide_width, prs.slide_height
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
    p.font.name = "Arial"
    p.font.size = Pt(44)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

def build_ppt(
    content_items,
    widescreen=True,
    shrink_to_fit=True,
    blank_slide_on_empty=False
) -> bytes:
    prs = Presentation()
    if widescreen:
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
    for item in content_items:
        if blank_slide_on_empty and (item.strip() == ""):
            add_black_slide(prs)
        else:
            add_text_slide(prs, item, shrink_to_fit=shrink_to_fit)
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

# ---------- UI ----------
st.subheader("Forrás")
tab1, tab2 = st.tabs(["📄 Fájl feltöltés", "📂 Meglévő útvonal (haladóknak)"])

uploaded = None
with tab1:
    f = st.file_uploader("Válassz .txt fájlt", type=["txt", "srt"])
    if f is not None:
        data = f.read()
        text = read_text_multi_enc(data)
        uploaded = ("uploaded", text)

with tab2:
    p = st.text_input("Teljes elérési út (pl. C:\\mappa\\feliratok.txt vagy /home/user/feliratok.txt)")
    if p:
        try:
            data = Path(p).read_bytes()
            text = read_text_multi_enc(data)
            uploaded = ("path", text)
        except Exception as e:
            st.error(f"Nem tudtam beolvasni: {e}")

st.subheader("Beállítások")
col1, col2 = st.columns(2)
with col1:
    mode = st.selectbox("Felosztás módja", ["line", "para", "srt"], index=0)
    widescreen = st.checkbox("Widescreen 16:9", value=True)
with col2:
    blank_on_empty = st.checkbox("Üres sor -> üres dia (line mód)", value=True)
    shrink = st.checkbox("Hosszú sor zsugorítása (shrink-to-fit)", value=True)

if st.button("PPTX generálása", type="primary", use_container_width=True):
    if not uploaded:
        st.warning("Adj meg forrást (fájl vagy útvonal).")
    else:
        _, raw_text = uploaded
        items = parse_text(raw_text, mode=mode, preserve_blanks=(mode=="line" and blank_on_empty))
        pptx_bytes = build_ppt(items, widescreen=widescreen, shrink_to_fit=shrink, blank_slide_on_empty=blank_on_empty)
        st.success(f"Siker! {len(items)} dia készült.")
        st.download_button(
            "PPTX letöltése",
            data=pptx_bytes,
            file_name="slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
