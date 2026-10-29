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
Голос → текст: микрофон в боковой панели, распознавание через Google (нужен интернет).
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
LABELS = {
    "school": "Мектеп аты",
    "bgzho": "БҒӘЖО",
    "date": "Мерзімі",
    "subject": "Пәні",
    "cls": "Сыныбы",
    "teacher": "Пән мұғалімі",
    "topic": "Сабақтың тақырыбы",
    "goal": "Сабаққа қатысу мақсатым",
    "visual": "Көрнекілігі",
    "notes": "Өзіме түртіп алдым",
}
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


def target_options():
    opts = {f"f_{k}": LABELS[k] for k in KEYS}
    for i, t in enumerate(st.session_state.row_types):
        names = COLS_SPLIT if t == "split" else COLS_SPAN
        for j, n in enumerate(names):
            opts[cell_key(i, j)] = f"Кесте {i + 1}-қатар · {n}"
    return opts


# ================================================================ колбэки
def on_audio():
    ss = st.session_state
    rec = ss.get(f"audio_{ss.audio_n}")
    if rec is None:
        return
    text, err = transcribe(rec.getvalue(), ss.lang)
    if text:
        tk = ss.target
        ss[tk] = append_text(ss.get(tk, ""), text)
        ss.voice_msg = ("ok", f"«{text}» → {target_options().get(tk, tk)}")
    else:
        ss.voice_msg = ("err", err)
    ss.audio_n += 1  # новый пустой микрофон для следующей записи


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
        if ss.get("target", "").startswith(f"r{i}_"):
            ss.target = "f_school"


def clear_all():
    lang = st.session_state.lang
    s = fresh_state()
    s["lang"] = lang
    put_into_session(s)
    st.session_state.target = "f_school"


# ================================================================ интерфейс
def app():
    st.set_page_config(page_title="Өзара сабаққа қатысу журналы", page_icon="📘", layout="centered")
    st.markdown("""
    <style>
    .stTextInput input, .stTextArea textarea{font-family:"Times New Roman",Times,serif;font-size:18px}
    .jr-title{font-family:"Times New Roman",serif;text-align:center}
    .jr-title .t1{font-weight:700;font-size:18px}
    .jr-title .t2{font-weight:700;font-size:32px;line-height:1.3;margin:22px 0 8px}
    .jr-title .t3{font-weight:700;font-size:20px;margin-bottom:10px}
    .jr-h{font-family:"Times New Roman",serif;font-weight:700;text-align:center;font-size:20px;margin:6px 0}
    </style>""", unsafe_allow_html=True)

    ss = st.session_state
    if "init" not in ss:
        put_into_session(load_file())
        ss.audio_n = 0
        ss.target = "f_school"
        ss.init = True

    # ---------- боковая панель: микрофон ----------
    with st.sidebar:
        st.header("🎤 Дауыспен жазу")
        st.selectbox("Дыбыс тілі", list(LANGS), format_func=LANGS.get, key="lang")
        opts = target_options()
        if ss.target not in opts:
            ss.target = "f_school"
        st.selectbox("Қай жолға жазамыз?", list(opts), format_func=opts.get, key="target")
        st.audio_input("Микрофонды басып сөйлеңіз, біткен соң тоқтатыңыз",
                       key=f"audio_{ss.audio_n}", on_change=on_audio)
        msg = ss.pop("voice_msg", None)
        if msg:
            (st.success if msg[0] == "ok" else st.error)(msg[1])
        st.caption("Танылған мәтін таңдалған жолдың соңына қосылады. "
                   "Кейін оны пернетақтамен түзете аласыз. Интернет қажет.")

    # ---------- титул ----------
    st.markdown('<div class="jr-title"><div class="t1">Павлодар облысы білім беру басқармасының</div></div>',
                unsafe_allow_html=True)
    st.text_input("Мектеп аты", key="f_school")
    st.markdown('<div class="jr-title"><div class="t2">Өзара сабаққа<br>қатысу журналы</div>'
                '<div class="t3">2026 – 2027 оқу жылы</div></div>', unsafe_allow_html=True)
    _, right = st.columns([1, 1])
    right.text_input("БҒӘЖО", key="f_bgzho")

    st.divider()

    # ---------- лист посещения ----------
    st.markdown('<div class="jr-h">Өзара сабаққа қатынасу парағы</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([2, 3, 1.3])
    c1.text_input("Мерзімі", key="f_date")
    c2.text_input("Пәні", key="f_subject")
    c3.text_input("Сыныбы", key="f_cls")
    st.text_input("Пән мұғалімі", key="f_teacher")
    st.text_input("Сабақтың тақырыбы", key="f_topic")
    st.text_area("Сабаққа қатысу мақсатым:", key="f_goal", height=120)
    st.text_input("Көрнекілігі:", key="f_visual")

    st.markdown("#### Кесте")
    for i, t in enumerate(ss.row_types):
        with st.container(border=True):
            st.caption(f"{i + 1}-қатар" + (" (Іс-әрекеттер екі бағанға бөлінген)" if t == "split" else ""))
            if t == "split":
                cols = st.columns([1.2, 3.5, 3.2, 2.1])
                names = COLS_SPLIT
            else:
                cols = st.columns([1.2, 6.7, 2.1])
                names = COLS_SPAN
            for j, (col, n) in enumerate(zip(cols, names)):
                col.text_area(n, key=cell_key(i, j), height=160 if i < 3 else 120)
    b1, b2, b3 = st.columns(3)
    b1.button("＋ Қатар қосу", on_click=add_row, args=("span",), use_container_width=True)
    b2.button("＋ Бөлінген қатар", on_click=add_row, args=("split",), use_container_width=True)
    b3.button("− Соңғы қатарды өшіру", on_click=del_row, use_container_width=True)

    st.text_area("Өзіме түртіп алдым:", key="f_notes", height=220)

    # ---------- сохранение ----------
    s = collect_state()
    try:
        save_file(s)
        saved_ok = True
    except OSError:
        saved_ok = False

    st.divider()
    d1, d2 = st.columns([2, 1])
    d1.download_button("📄 Word-қа сақтау (.docx)", data=build_docx(s), file_name=DOCX_NAME,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       type="primary", use_container_width=True)
    with d2.popover("Тазалау", use_container_width=True):
        st.write("Барлық жазылған мәтін өшіріледі.")
        st.button("Иә, бәрін өшіру", on_click=clear_all, type="primary")
    st.caption(("✓ Автоматты сақталды " + datetime.now().strftime("%H:%M:%S")) if saved_ok
               else "⚠ Файлға сақтау мүмкін болмады")


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
