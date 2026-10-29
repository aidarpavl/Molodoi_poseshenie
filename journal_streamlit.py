#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Өзара сабаққа қатысу журналы — Streamlit-версия.

Установка (один раз):
    pip install streamlit SpeechRecognition

Запуск:
    streamlit run journal_streamlit.py
    (или просто: python journal_streamlit.py — он сам запустит Streamlit)

Записи сохраняются в journal_data.json рядом со скриптом.
Голос → текст: кнопка 🎤 рядом с каждым полем, распознавание через Google (нужен интернет).
"""

import io
import json
import os
import sys
import zipfile
from datetime import datetime

import streamlit as st

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "journal_data.json")
DOCX_NAME = "Өзара_сабаққа_қатысу_журналы.docx"

KEYS = ["school", "bgzho", "date", "subject", "cls", "teacher", "topic", "goal", "visual", "notes"]
COLS_SPAN = ["Сабақ кезеңдері", "Іс - әрекеттер", "Әдіс - тәсілдер"]
COLS_SPLIT = ["Сабақ кезеңдері", "Іс - әрекеттер (1)", "Іс - әрекеттер (2)", "Әдіс - тәсілдер"]
LANGS = {"kk-KZ": "Қазақша", "ru-RU": "Русский"}


# ================================================================ данные
def fresh_state():
    return {
        "F": {k: "" for k in KEYS},
        "rows": [
            {"t": "span", "c": ["", "", ""]},
            {"t": "span", "c": ["", "", ""]},
            {"t": "split", "c": ["", "", "", ""]},
        ],
        "lang": "kk-KZ",
    }


def clean_state(d):
    s = fresh_state()
    if not isinstance(d, dict):
        return s
    f = d.get("F")
    if isinstance(f, dict):
        for k in KEYS:
            s["F"][k] = f.get(k) if isinstance(f.get(k), str) else ""
    rows = d.get("rows")
    if isinstance(rows, list):
        out = []
        for r in rows:
            if not isinstance(r, dict):
                continue
            t = "split" if r.get("t") == "split" else "span"
            n = 4 if t == "split" else 3
            c = r.get("c") if isinstance(r.get("c"), list) else []
            c = [(x if isinstance(x, str) else "") for x in c][:n]
            out.append({"t": t, "c": c + [""] * (n - len(c))})
        if out:
            s["rows"] = out
    if d.get("lang") in LANGS:
        s["lang"] = d["lang"]
    return s


def load_file():
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as fh:
            return clean_state(json.load(fh))
    except (OSError, ValueError):
        return fresh_state()


def save_file(s):
    tmp = DATA_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(s, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, DATA_FILE)


def cell_key(r, c):
    return f"r{r}_c{c}"


def put_into_session(s):
    """Записывает состояние в session_state (ключи виджетов)."""
    ss = st.session_state
    for k in KEYS:
        ss[f"f_{k}"] = s["F"][k]
    # убрать старые ключи ячеек
    for k in [k for k in ss.keys() if isinstance(k, str) and k.startswith("r") and "_c" in k]:
        del ss[k]
    ss.row_types = [r["t"] for r in s["rows"]]
    for i, r in enumerate(s["rows"]):
        for j, v in enumerate(r["c"]):
            ss[cell_key(i, j)] = v
    ss.lang = s["lang"]


def collect_state():
    ss = st.session_state
    rows = []
    for i, t in enumerate(ss.row_types):
        n = 4 if t == "split" else 3
        rows.append({"t": t, "c": [ss.get(cell_key(i, j), "") for j in range(n)]})
    return {"F": {k: ss.get(f"f_{k}", "") for k in KEYS}, "rows": rows, "lang": ss.get("lang", "kk-KZ")}


# ================================================================ DOCX
def xe(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


RFN = ('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
       'w:cs="Times New Roman" w:eastAsia="Times New Roman"/>')


def w_run(t, b=False, u=False, sz=24):
    rp = (RFN + ("<w:b/>" if b else "") + ('<w:u w:val="single"/>' if u else "")
          + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:lang w:val="kk-KZ"/>')
    return "".join(
        f'<w:r><w:rPr>{rp}</w:rPr>{"<w:br/>" if i else ""}<w:t xml:space="preserve">{xe(line)}</w:t></w:r>'
        for i, line in enumerate(str(t).split("\n")))


def w_para(inner="", pb=False, before=0, after=0, ind=None, jc=None):
    pp = ("<w:pageBreakBefore/>" if pb else "") \
        + f'<w:spacing w:before="{before}" w:after="{after}" w:line="240" w:lineRule="auto"/>' \
        + (f'<w:ind w:left="{ind}"/>' if ind else "") + (f'<w:jc w:val="{jc}"/>' if jc else "")
    return f"<w:p><w:pPr>{pp}</w:pPr>{inner}</w:p>"


def w_val(x, n, **o):
    x = (x or "").strip()
    return w_run(x, u=True, **o) if x else w_run("_" * n, **o)


def w_cell(w, text, span=0, jc=None):
    bd = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                 for s in ("top", "left", "bottom", "right"))
    gs = f'<w:gridSpan w:val="{span}"/>' if span else ""
    return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{gs}<w:tcBorders>{bd}</w:tcBorders></w:tcPr>'
            f'{w_para(w_run(text or ""), jc=jc)}</w:tc>')


def docx_table(s):
    hdr = ('<w:tr><w:trPr><w:trHeight w:val="800" w:hRule="atLeast"/><w:tblHeader/></w:trPr>'
           + w_cell(1277, "Сабақ кезеңдері", 0, "center") + w_cell(7087, "Іс - әрекеттер", 2, "center")
           + w_cell(2268, "Әдіс - тәсілдер", 0, "center") + "</w:tr>")
    rows = []
    for i, r in enumerate(s["rows"]):
        h = 2500 if i == 0 else 6000 if i < 3 else 3000
        c = r["c"]
        if r["t"] == "split":
            cells = w_cell(1277, c[0]) + w_cell(3685, c[1]) + w_cell(3402, c[2]) + w_cell(2268, c[3])
        else:
            cells = w_cell(1277, c[0]) + w_cell(7087, c[1], 2) + w_cell(2268, c[2])
        rows.append(f'<w:tr><w:trPr><w:trHeight w:val="{h}" w:hRule="atLeast"/></w:trPr>{cells}</w:tr>')
    return ('<w:tbl><w:tblPr><w:tblW w:w="10632" w:type="dxa"/><w:tblInd w:w="-743" w:type="dxa"/>'
            '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/>'
            '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'
            '<w:gridCol w:w="1277"/><w:gridCol w:w="3685"/><w:gridCol w:w="3402"/><w:gridCol w:w="2268"/>'
            f'</w:tblGrid>{hdr}{"".join(rows)}</w:tbl>')


def docx_body(s):
    F, ind = s["F"], 567
    b = w_para(w_run("Павлодар облысы білім беру басқармасының", b=True), jc="center", after=120)
    b += w_para(w_run("мектеп аты ", b=True) + w_val(F["school"], 28, b=True), jc="center")
    b += w_para() * 5
    b += w_para(w_run("Өзара сабаққа", b=True, sz=40), jc="center", after=120)
    b += w_para(w_run("қатысу журналы", b=True, sz=40), jc="center")
    b += w_para() * 2
    b += w_para(w_run("2026 – 2027 оқу жылы", b=True, sz=28), jc="center")
    b += w_para(w_run("БҒӘЖО: ", b=True, sz=20) + w_val(F["bgzho"], 20, b=True, sz=20), jc="right", before=9800)
    b += w_para(w_run("Өзара сабаққа қатынасу парағы", b=True), jc="center", after=200, pb=True)
    b += w_para(w_run("Мерзімі ") + w_val(F["date"], 14) + w_run("      Пәні  ") + w_val(F["subject"], 19)
                + w_run("   Сыныбы ") + w_val(F["cls"], 5), ind=ind, after=160)
    b += w_para(w_run("Пән мұғалімі ") + w_val(F["teacher"], 60), ind=ind, after=160)
    b += w_para(w_run("Сабақтың тақырыбы ") + w_val(F["topic"], 48), ind=ind, after=160)
    b += w_para(w_run("Сабаққа қатысу мақсатым: ") + w_val(F["goal"], 270), ind=ind, after=160)
    b += w_para(w_run("Көрнекілігі: ") + w_val(F["visual"], 62), ind=ind, after=160)
    b += w_para() * 2
    b += docx_table(s)
    b += w_para(w_run("Өзіме түртіп алдым: ") + w_val(F["notes"], 420), before=60)
    b += ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="567" w:right="850" w:bottom="1134" '
          'w:left="1701" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>')
    return b


def build_docx(s):
    W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    files = {
        "[Content_Types].xml": H + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
        "_rels/.rels": H + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": H + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/></Relationships>',
        "word/styles.xml": H + f'<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>{RFN}'
        '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="kk-KZ"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>',
        "word/document.xml": H + f"<w:document {W}><w:body>{docx_body(s)}</w:body></w:document>",
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data.encode("utf-8"))
    return buf.getvalue()


# ================================================================ голос → текст
def transcribe(wav_bytes, lang):
    try:
        import speech_recognition as sr
    except ImportError:
        return None, "SpeechRecognition кітапханасы орнатылмаған: pip install SpeechRecognition"
    r = sr.Recognizer()
    try:
        with sr.AudioFile(io.BytesIO(wav_bytes)) as src:
            audio = r.record(src)
        return r.recognize_google(audio, language=lang), None
    except sr.UnknownValueError:
        return None, "Сөз танылмады. Анығырақ, жақынырақ сөйлеп көріңіз."
    except sr.RequestError:
        return None, "Тану қызметіне қосылу мүмкін болмады. Интернетті тексеріңіз."
    except Exception as e:  # повреждённый файл и т.п.
        return None, f"Қате: {e}"


def append_text(old, new):
    new = new.strip()
    if not new:
        return old
    if old.strip() == "" or old.rstrip().endswith((".", "!", "?")):
        new = new[0].upper() + new[1:]
    sep = "" if (not old or old.endswith((" ", "\n"))) else " "
    return old + sep + new


# ================================================================ колбэки
def on_audio(target):
    """Распознаёт запись из микрофона рядом с полем и дописывает текст в это поле."""
    ss = st.session_state
    n = ss.mic_n.get(target, 0)
    rec = ss.get(f"au_{target}_{n}")
    if rec is None:
        return
    text, err = transcribe(rec.getvalue(), ss.lang)
    if text:
        ss[target] = append_text(ss.get(target, ""), text)
        ss.voice_msg = ("✅", f"«{text}»")
    else:
        ss.voice_msg = ("⚠️", err)
    ss.mic_n[target] = n + 1  # новый пустой рекордер для следующей записи


def add_row(kind):
    ss = st.session_state
    i = len(ss.row_types)
    ss.row_types.append(kind)
    for j in range(4 if kind == "split" else 3):
        ss[cell_key(i, j)] = ""


def del_row():
    ss = st.session_state
    if len(ss.row_types) > 1:
        i = len(ss.row_types) - 1
        for j in range(4):
            ss.pop(cell_key(i, j), None)
        ss.row_types.pop()


def clear_all():
    lang = st.session_state.lang
    s = fresh_state()
    s["lang"] = lang
    put_into_session(s)


# ================================================================ виджеты
def mic(container, key):
    """Кнопка 🎤 рядом с полем: открывает рекордер, текст попадает в это поле."""
    n = st.session_state.mic_n.get(key, 0)
    with container.popover("🎤", help="Микрофон арқылы жазу"):
        st.caption("Басып сөйлеңіз, біткен соң тоқтатыңыз")
        st.audio_input("Микрофон", key=f"au_{key}_{n}", on_change=on_audio, args=(key,),
                       label_visibility="collapsed")


def field(label, key, area=False, height=None, container=None):
    box = container or st
    c1, c2 = box.columns([12, 1.3], vertical_alignment="bottom", gap="small")
    if area:
        c1.text_area(label, key=key, height=height)
    else:
        c1.text_input(label, key=key)
    mic(c2, key)


CSS = """
<style>
.stApp{background:#e9ebf1}
.block-container,[data-testid="stMainBlockContainer"]{background:#fff;max-width:900px;
  padding:2.2rem 2.6rem 3rem;margin-top:1.2rem;margin-bottom:2rem;border-radius:4px;
  box-shadow:0 2px 14px rgba(34,36,43,.18)}
.stTextInput input,.stTextArea textarea{font-family:"Times New Roman",Times,serif;font-size:18px;color:#050038}
.stTextInput label p,.stTextArea label p{font-family:"Times New Roman",Times,serif;font-size:17px;color:#1a1a2e}
.jr{font-family:"Times New Roman",Times,serif;text-align:center;color:#1a1a2e}
.jr .t1{font-weight:700;font-size:19px}
.jr .t2{font-weight:700;font-size:34px;line-height:1.35;margin:46px 0 14px}
.jr .t3{font-weight:700;font-size:22px;margin-bottom:30px}
.jr-h{font-family:"Times New Roman",Times,serif;font-weight:700;text-align:center;font-size:20px;margin:4px 0 8px}
.tip{background:#fff8d6;border:1px solid #f0dc7a;border-radius:8px;padding:10px 14px;font-size:14px;line-height:1.45;margin-bottom:10px}
.th{font-family:"Times New Roman",Times,serif;text-align:center;font-size:16px;border:1px solid #000;padding:6px 2px;margin-bottom:4px}
@media (max-width:640px){.block-container,[data-testid="stMainBlockContainer"]{padding:1.2rem 1rem 2rem}
  .jr .t2{font-size:26px;margin:26px 0 10px}}
</style>
"""


# ================================================================ интерфейс
def app():
    st.set_page_config(page_title="Өзара сабаққа қатысу журналы", page_icon="📘", layout="centered")
    st.markdown(CSS, unsafe_allow_html=True)

    ss = st.session_state
    if "init" not in ss:
        put_into_session(load_file())
        ss.mic_n = {}
        ss.init = True

    msg = ss.pop("voice_msg", None)
    if msg:
        st.toast(msg[1], icon=msg[0])

    # ---------- верхняя панель ----------
    h1, h2 = st.columns([3, 1.4], vertical_alignment="bottom")
    h1.markdown("### 📘 Өзара сабаққа қатысу журналы")
    h2.selectbox("Дыбыс тілі", list(LANGS), format_func=LANGS.get, key="lang")
    st.markdown('<div class="tip">🎤 Кез келген жолдың немесе кесте ұяшығының жанындағы микрофон '
                'батырмасын басып сөйлеңіз — дауыс мәтінге айналып, сол жолға қосылады. '
                'Жазылған мәтінді кейін <b>пернетақтамен</b> түзетіп немесе толықтыра аласыз. '
                'Бәрі автоматты түрде сақталып отырады. (Дауысты тану үшін интернет қажет.)</div>',
                unsafe_allow_html=True)

    # ---------- 1-бет: мұқаба ----------
    st.markdown('<div class="jr"><div class="t1">Павлодар облысы білім беру басқармасының</div></div>',
                unsafe_allow_html=True)
    _, mid, _ = st.columns([1, 3, 1])
    field("мектеп аты", "f_school", container=mid)
    st.markdown('<div class="jr"><div class="t2">Өзара сабаққа<br>қатысу журналы</div>'
                '<div class="t3">2026 – 2027 оқу жылы</div></div>', unsafe_allow_html=True)
    _, right = st.columns([1.4, 1])
    field("БҒӘЖО:", "f_bgzho", container=right)

    st.divider()

    # ---------- 2-бет: қатысу парағы ----------
    st.markdown('<div class="jr-h">Өзара сабаққа қатынасу парағы</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([2, 3, 1.6])
    field("Мерзімі", "f_date", container=c1)
    field("Пәні", "f_subject", container=c2)
    field("Сыныбы", "f_cls", container=c3)
    field("Пән мұғалімі", "f_teacher")
    field("Сабақтың тақырыбы", "f_topic")
    field("Сабаққа қатысу мақсатым:", "f_goal", area=True, height=130)
    field("Көрнекілігі:", "f_visual")

    # ---------- кесте ----------
    st.write("")
    hc = st.columns([1.2, 6.7, 2.1], gap="small")
    for col, name in zip(hc, ["Сабақ кезеңдері", "Іс - әрекеттер", "Әдіс - тәсілдер"]):
        col.markdown(f'<div class="th">{name}</div>', unsafe_allow_html=True)
    for i, t in enumerate(ss.row_types):
        with st.container(border=True):
            if t == "split":
                cols = st.columns([1.2, 3.5, 3.2, 2.1], gap="small")
                names = COLS_SPLIT
            else:
                cols = st.columns([1.2, 6.7, 2.1], gap="small")
                names = COLS_SPAN
            for j, (col, n) in enumerate(zip(cols, names)):
                k = cell_key(i, j)
                col.text_area(n, key=k, height=170 if i < 3 else 120, label_visibility="collapsed",
                              placeholder=n)
                mic(col, k)
    b1, b2, b3 = st.columns(3)
    b1.button("＋ Қатар қосу", on_click=add_row, args=("span",))
    b2.button("＋ Бөлінген қатар", on_click=add_row, args=("split",))
    b3.button("− Соңғы қатарды өшіру", on_click=del_row)

    field("Өзіме түртіп алдым:", "f_notes", area=True, height=240)

    # ---------- сохранение ----------
    s = collect_state()
    try:
        save_file(s)
        saved = "✓ Автоматты сақталды " + datetime.now().strftime("%H:%M:%S")
    except OSError:
        saved = "⚠ Файлға сақтау мүмкін болмады"

    st.divider()
    d1, d2, d3 = st.columns([2.2, 1, 2], vertical_alignment="center")
    d1.download_button("📄 Word-қа сақтау (.docx)", data=build_docx(s), file_name=DOCX_NAME,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       type="primary")
    with d2.popover("Тазалау"):
        st.write("Барлық жазылған мәтін өшіріледі. Жалғастырасыз ба?")
        st.button("Иә, бәрін өшіру", on_click=clear_all, type="primary")
    d3.caption(saved)


# ================================================================ запуск
def _inside_streamlit():
    try:
        from streamlit.runtime import exists
        return exists()
    except Exception:
        return False


if _inside_streamlit():
    app()
elif __name__ == "__main__":
    # «python journal_streamlit.py» → перезапуск через «streamlit run»
    from streamlit.web import cli as stcli
    sys.argv = ["streamlit", "run", os.path.abspath(__file__)]
    sys.exit(stcli.main())
