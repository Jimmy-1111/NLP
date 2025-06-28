# -*- coding: utf-8 -*-
import streamlit as st
import pdfplumber, pandas as pd, re, io
import tempfile, subprocess, os
from pathlib import Path

# ---------- 抽取層 ---------- #
def has_enough_cjk(s, thresh=0.1):
    cjk = re.findall(r'[\u3040-\u30ff\u4e00-\u9fff]', s)
    return len(cjk) / max(len(s), 1) >= thresh

def extract_pdfplumber(page):
    return page.extract_text() or ""

def extract_poppler(pdf_bytes, page_no):
    # 寫進暫存 PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tf:
        tf.write(pdf_bytes)
        pdf_path = tf.name
    txt_path = pdf_path.replace(".pdf", ".txt")
    try:
        subprocess.run(
            ["pdftotext", "-layout", "-enc", "UTF-8",
             "-f", str(page_no), "-l", str(page_no),
             pdf_path, txt_path],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
        )
        txt = Path(txt_path).read_text(encoding="utf-8", errors="ignore")
        return txt
    except Exception:
        return ""
    finally:
        for p in (pdf_path, txt_path):
            if os.path.exists(p): os.unlink(p)

def extract_ocr(pdf_bytes, page_no):
    try:
        from pdf2image import convert_from_bytes
        import pytesseract
        img = convert_from_bytes(pdf_bytes, dpi=300,
                                 first_page=page_no, last_page=page_no)[0]
        return pytesseract.image_to_string(img, lang="jpn")
    except Exception:
        return ""

def extract_page_text(page, pdf_bytes, page_no, debug=False):
    # 1. pdfplumber
    txt = extract_pdfplumber(page)
    if debug: st.write(f"📄{page_no} pdfplumber 字數: {len(txt)}  CJK 比: "
                       f"{len(re.findall(r'[\\u3040-\\u30ff\\u4e00-\\u9fff]', txt))/max(len(txt),1):.1%}")
    if has_enough_cjk(txt): return txt

    # 2. Poppler
    txt = extract_poppler(pdf_bytes, page_no)
    if debug: st.write(f"📄{page_no} pdftotext  字數: {len(txt)}  CJK 比: "
                       f"{len(re.findall(r'[\\u3040-\\u30ff\\u4e00-\\u9fff]', txt))/max(len(txt),1):.1%}")
    if has_enough_cjk(txt): return txt

    # 3. OCR
    txt = extract_ocr(pdf_bytes, page_no)
    if debug: st.write(f"📄{page_no} OCR        字數: {len(txt)}  CJK 比: "
                       f"{len(re.findall(r'[\\u3040-\\u30ff\\u4e00-\\u9fff]', txt))/max(len(txt),1):.1%}")
    return txt


# ---------- 分句 ---------- #
def split_sentences(text):
    if not text:
        return []
    cleaned = " ".join(l.strip() for l in text.splitlines() if l.strip())
    raw = re.split(r'(?<=[。．！？.!?])', cleaned)

    exclude_keywords = [
        "EDINET提出書類", "有価証券報告書",
        re.compile(r'.+株式会社\(E\d{5}\)'), re.compile(r'^\d{1,3}/\d{1,3}$')
    ]

    sents = []
    for s in raw:
        s = s.strip()
        if not s or len(s) < 5 or re.fullmatch(r'^[\d\W\s]+$', s):
            continue
        if not re.search(r'[\u3040-\u30ff\u4e00-\u9fff]', s):
            continue
        if any((kw in s) if isinstance(kw, str) else kw.search(s)
               for kw in exclude_keywords):
            continue
        sents.append(s)
    return sents


# ---------- Streamlit 主體 ---------- #
def main():
    st.title("📄 PDF 分句下載器（plumber → Poppler → OCR）")
    st.markdown("**依序嘗試 pdfplumber → pdftotext → Tesseract-OCR**，抓到日文即分句。")

    # ----- 範圍選擇 ----- #
    if "ranges" not in st.session_state: st.session_state["ranges"] = []

    pdf_file = st.file_uploader("上傳 PDF", type="pdf")
    c1, c2 = st.columns(2)
    with c1:
        sp = st.number_input("開始頁", 1, step=1, value=1)
    with c2:
        ep = st.number_input("結束頁", 1, step=1, value=1)
    if st.button("➕ 新增範圍"):
        if ep >= sp:
            st.session_state["ranges"].append((sp, ep))
        else:
            st.warning("結束頁必須 ≥ 開始頁")
    for i, (s, e) in enumerate(st.session_state["ranges"], 1):
        st.write(f"{i}. 第 {s}–{e} 頁")

    # ----- 命名欄 ----- #
    company = st.text_input("企業名稱")
    year    = st.text_input("年份(4位)")
    month   = st.text_input("月")
    day     = st.text_input("日")

    # ----- 執行 ----- #
    if st.button("🚀 開始處理") and pdf_file and st.session_state["ranges"]:
        st.info("解析中…")
        out_name = "_".join(filter(None, [company, year, month, day])) or "output"
        out_name += ".xlsx"

        bytes_data = pdf_file.getvalue()
        data = []
        with pdfplumber.open(io.BytesIO(bytes_data)) as pdf:
            total_pages = sum(e - s + 1 for s, e in st.session_state["ranges"])
            prog = st.progress(0)
            done = 0
            for s, e in st.session_state["ranges"]:
                for idx_page in range(s - 1, e):
                    if idx_page >= len(pdf.pages): continue
                    page = pdf.pages[idx_page]
                    text = extract_page_text(page, bytes_data, idx_page + 1, debug=True)
                    for idx_sent, sent in enumerate(split_sentences(text), 1):
                        data.append({"頁碼": idx_page + 1,
                                     "語句編號": idx_sent,
                                     "語句內容": sent})
                    done += 1
                    prog.progress(done / total_pages)

        df = pd.DataFrame(data)
        if df.empty:
            st.error("⚠️ 三層抽取皆未抓到有效日文，或全部句子被過濾。可檢查 Poppler/Tesseract 安裝情況，或放寬過濾條件。")
            return
        st.success("完成！")
        st.dataframe(df, use_container_width=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False)
        st.download_button("📥 下載 Excel", buf.getvalue(), file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.session_state["ranges"] = []  # reset


if __name__ == "__main__":
    main()
