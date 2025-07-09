import streamlit as st
import pdfplumber
import pandas as pd
import pytesseract
from PIL import Image
import re
import io
import os

# 設定 pytesseract 執行檔路徑（macOS Homebrew 安裝預設路徑）
pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"
os.environ["TESSDATA_PREFIX"] = "/opt/homebrew/share/"

def split_sentences(text):
    if not text:
        return []

    lines = text.splitlines()
    merged_lines = []
    buffer = ""

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if re.match(r'^\d{1,2}[ 　]*【.+】$', line) or re.match(r'^[(（][0-9０-９]{1,3}[)）]', line):
            if buffer:
                merged_lines.append(buffer)
                buffer = ""
            merged_lines.append(line)
            continue

        if buffer:
            if (
                not re.search(r'[。．！？.!?]$', buffer)
                and re.match(r'^[ぁ-んァ-ンa-zＡ-Ｚａ-ｚ一-龯A-Z（(【「『]', line)
            ):
                buffer += line
            else:
                merged_lines.append(buffer)
                buffer = line
        else:
            buffer = line

    if buffer:
        merged_lines.append(buffer)

    split_by_punctuation = []
    for line in merged_lines:
        if re.match(r'^[(（]?[0-9０-９]{1,3}[)）]?.*$', line):
            split_by_punctuation.append(line)
        else:
            segments = re.split(r'([。．！？.!?])', line)
            sentence = ""
            for seg in segments:
                if seg in "。．！？.!?":
                    sentence += seg
                    if sentence.strip():
                        split_by_punctuation.append(sentence.strip())
                        sentence = ""
                else:
                    sentence += seg
            if sentence.strip():
                split_by_punctuation.append(sentence.strip())

    exclude_keywords = [
        "EDINET提出書類",
        "有価証券報告書",
        re.compile(r'.+株式会社\(E\d{5}\)'),
        re.compile(r'^\d{1,3}/\d{1,3}$')
    ]

    sentences = []
    for s in split_by_punctuation:
        exclude = False
        for kw in exclude_keywords:
            if isinstance(kw, str) and kw in s:
                exclude = True
                break
            elif isinstance(kw, re.Pattern) and kw.search(s):
                exclude = True
                break
        if not exclude:
            sentences.append(s)

    return sentences

def main():
    st.title("📄 PDF/圖片語句分割器")
    st.write("上傳 PDF 或圖片進行語句分割，支援日文 OCR 與分句")

    if "ranges" not in st.session_state:
        st.session_state["ranges"] = []

    # PDF 區段
    st.subheader("📕 PDF 分句區段")
    pdf_file = st.file_uploader("請上傳 PDF 檔案", type="pdf")

    start_page = st.number_input("開始頁碼（從 1 起算）", min_value=1, step=1, key="start_page")
    end_page = st.number_input("結束頁碼（包含）", min_value=1, step=1, key="end_page")

    if st.button("➕ 新增頁碼範圍"):
        if end_page >= start_page:
            st.session_state["ranges"].append((start_page, end_page))
        else:
            st.warning("⚠️ 結束頁碼不得小於開始頁碼")

    if st.session_state["ranges"]:
        st.markdown("🗂️ **已選範圍：**")
        for idx, (s, e) in enumerate(st.session_state["ranges"]):
            st.write(f"{idx+1}. 第 {s} 到第 {e} 頁")

    # 圖片 OCR 區段
    st.subheader("🖼️ 圖片 OCR 區段")
    image_file = st.file_uploader("上傳圖片（JPG、PNG）進行 OCR 與分句", type=["png", "jpg", "jpeg"])

    # 基本欄位輸入
    st.subheader("📝 基本資料與下載命名")
    company = st.text_input("企業名稱（中日英文、數字、日文假名、符號皆可）")
    year = st.text_input("年份（例如：2024）")
    month = st.text_input("月份（可空白）")
    day = st.text_input("日期（可空白）")
    custom_filename = st.text_input("（選填）自訂檔名（含 .xlsx 或不含皆可）", "")

    valid_company = bool(re.match(r"^[\u4e00-\u9fa5A-Za-z0-9\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uFF66-\uFF9D\u3000-\u303F・ー\s\-\(\)\[\]【】『』「」、。]+$", company))
    valid_year = bool(re.match(r"^\d{4}$", year))
    valid_month = (month == '' or re.match(r"^(0?[1-9]|1[0-2])$", month))
    valid_day = (day == '' or re.match(r"^(0?[1-9]|[12][0-9]|3[01])$", day))

    if company and not valid_company:
        st.error("❌ 企業名稱只能包含中日英文、數字、假名與常用符號。")
    if year and not valid_year:
        st.error("❌ 年份必須是4位數字")
    if month and not valid_month:
        st.error("❌ 月份格式錯誤（請輸入 1~12 或留空）")
    if day and not valid_day:
        st.error("❌ 日期格式錯誤（請輸入 1~31 或留空）")

    # 檔案命名
    if custom_filename:
        filename = custom_filename if custom_filename.endswith(".xlsx") else custom_filename + ".xlsx"
    else:
        filename_parts = [company.strip(), year.strip()]
        if month:
            filename_parts.append(str(int(month)))
        if day:
            filename_parts.append(str(int(day)))
        filename = "_".join(filename_parts) + ".xlsx"

    # 主按鈕
    if st.button("🚀 選擇結束，開始分割"):
        if not (company and year and valid_company and valid_year and valid_month and valid_day):
            st.warning("⚠️ 請正確填寫企業名稱與年份（必填），月份/日期可空白。")
            return

        data = []

        if pdf_file and st.session_state["ranges"]:
            with pdfplumber.open(pdf_file) as pdf:
                for (start_page, end_page) in st.session_state["ranges"]:
                    for i in range(start_page - 1, end_page):
                        if i < len(pdf.pages):
                            page = pdf.pages[i]
                            text = page.extract_text()
                            sentences = split_sentences(text)
                            if sentences and re.match(r'^\d{1,3}/\d{1,3}$', sentences[0]):
                                sentences = sentences[1:]
                            for idx, s in enumerate(sentences, 1):
                                data.append({"來源": f"PDF第{i+1}頁", "語句編號": idx, "語句內容": s})

        if image_file:
            image = Image.open(image_file)
            ocr_text = pytesseract.image_to_string(image, lang="jpn")
            sentences = split_sentences(ocr_text)
            for idx, s in enumerate(sentences, 1):
                data.append({"來源": "圖片OCR", "語句編號": idx, "語句內容": s})

        if data:
            df = pd.DataFrame(data)
            st.success("✅ 處理完成！")
            st.dataframe(df, use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            st.download_button(
                label="📥 下載 Excel 檔",
                data=output.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.session_state["ranges"] = []
        else:
            st.warning("⚠️ 沒有處理任何資料，請確認有上傳 PDF 或圖片。")

if __name__ == '__main__':
    main()
