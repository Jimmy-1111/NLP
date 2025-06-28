import streamlit as st
import pdfplumber
import pandas as pd
import re, io, subprocess, tempfile, os, sys
from pathlib import Path

# ---------- 文字抽取核心 ---------- #
def extract_page_text(page, page_number, pdf_path=None):
    """
    優先順序：
    1) pdfplumber.extract_text()
    2) Poppler pdftotext (需 pdf_path)
    3) OCR (pdf2image + pytesseract)
    """
    # --- 1. pdfplumber ---
    txt = page.extract_text() or ""

    if has_enough_cjk(txt):
        return txt

    # --- 2. pdftotext ---
    if pdf_path:
        try:
            with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tf:
                tf_path = tf.name
            # -layout 讓排版接近原樣；-f/-l 指定頁碼
            subprocess.run(
                ["pdftotext", "-layout", "-enc", "UTF-8",
                 "-f", str(page_number), "-l", str(page_number),
                 pdf_path, tf_path],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            with open(tf_path, "r", encoding="utf-8") as f:
                txt2 = f.read()
            os.unlink(tf_path)
            if has_enough_cjk(txt2):
                return txt2
        except Exception as e:
            # pdftotext 失敗就往下
            pass

    # --- 3. OCR pytesseract ---
    try:
        from pdf2image import convert_from_bytes
        import pytesseract
        img = convert_from_bytes(page.pdf.within_pdf.stream.get_data(),
                                 dpi=300, first_page=1, last_page=1)[0]
        txt3 = pytesseract.image_to_string(img, lang="jpn")
        return txt3
    except Exception as e:
        return txt  # 至少回傳原始結果


def has_enough_cjk(s, thresh=0.1):
    """判斷字串中日文/中文比例是否足夠"""
    if not s:
        return False
    cjk_chars = re.findall(r'[\u3040-\u30ff\u4e00-\u9fff]', s)
    return len(cjk_chars) / len(s) >= thresh


# ---------- 分句 ---------- #
def split_sentences(text):
    if not text:
        return []

    cleaned_text = " ".join([line.strip() for line in text.splitlines() if line.strip()])
    raw_sentences = re.split(r'(?<=[。．！？.!?])', cleaned_text)

    exclude_keywords = [
        "EDINET提出書類",
        "有価証券報告書",
        re.compile(r'.+株式会社\(E\d{5}\)'),
        re.compile(r'^\d{1,3}/\d{1,3}$')
    ]

    sentences = []
    for s in raw_sentences:
        s = s.strip()
        if not s:
            continue
        if len(s) < 5 or re.fullmatch(r'^[\d\W\s]+$', s):
            continue
        # 至少有一個 CJK 字
        if not re.search(r'[\u3040-\u30ff\u4e00-\u9fff]', s):
            continue

        if any((kw in s) if isinstance(kw, str) else kw.search(s) for kw in exclude_keywords):
            continue
        sentences.append(s)
    return sentences


# ---------- Streamlit 主體 ---------- #
def main():
    st.title("📄 PDF 語句分割器（含自動 OCR / Poppler）")
    st.write("上傳 PDF，選頁碼範圍，自動偵測最佳抽取方式，分句後下載 Excel。")

    if "ranges" not in st.session_state:
        st.session_state["ranges"] = []

    pdf_file = st.file_uploader("請上傳 PDF 檔案", type="pdf")

    start_page = st.number_input("開始頁碼（從 1 起算）", min_value=1, step=1)
    end_page = st.number_input("結束頁碼（包含）", min_value=1, step=1)

    if st.button("➕ 新增範圍"):
        if end_page >= start_page:
            st.session_state["ranges"].append((start_page, end_page))
        else:
            st.warning("⚠️ 結束頁碼不得小於開始頁碼")

    if st.session_state["ranges"]:
        st.markdown("🗂️ **已選範圍：**")
        for idx, (s, e) in enumerate(st.session_state["ranges"]):
            st.write(f"{idx+1}. 第 {s}–{e} 頁")

    company = st.text_input("企業名稱")
    year = st.text_input("年份（4 位數）")
    month = st.text_input("月份（可空白）")
    day = st.text_input("日期（可空白）")

    # 檔名
    filename = "_".join(filter(None, [company.strip(), year.strip(), month.strip(), day.strip()])) + ".xlsx"

    if st.button("🚀 開始處理") and pdf_file and st.session_state["ranges"]:
        st.info("⏳ 正在解析 PDF，請稍候…")
        data = []
        # 將上傳的 BytesIO 寫入臨時檔，以便 pdftotext 使用
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(pdf_file.read())
            tmp_path = tmp_pdf.name

        with pdfplumber.open(tmp_path) as pdf:
            progress = st.progress(0)
            total_pages = sum(e - s + 1 for s, e in st.session_state["ranges"])
            processed = 0

            for (s_page, e_page) in st.session_state["ranges"]:
                for i in range(s_page - 1, e_page):
                    if i >= len(pdf.pages):
                        continue
                    page = pdf.pages[i]
                    text = extract_page_text(page, i + 1, pdf_path=tmp_path)
                    sentences = split_sentences(text)
                    for idx, s in enumerate(sentences, 1):
                        data.append({"頁碼": i + 1, "語句編號": idx, "語句內容": s})
                    processed += 1
                    progress.progress(processed / total_pages)

        df = pd.DataFrame(data)
        st.success("✅ 完成！")
        st.dataframe(df, use_container_width=True)

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 下載 Excel", data=out.getvalue(),
                           file_name=filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # 清理暫存 PDF
        Path(tmp_path).unlink(missing_ok=True)
        st.session_state["ranges"] = []


if __name__ == "__main__":
    main()
