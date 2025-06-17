import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

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

        # 主標或條列標題：保留獨立句
        if re.match(r'^\d{1,2}[ 　]*【.+】$', line) or re.match(r'^[(（][0-9０-９]{1,3}[)）]', line):
            if buffer:
                merged_lines.append(buffer)
                buffer = ""
            merged_lines.append(line)
            continue

        if buffer:
            if (
                not re.search(r'[。．！？]$', buffer)
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
            segments = re.split(r'(。)', line)
            sentence = ""
            for seg in segments:
                if seg == "。":
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
    st.title("PDF 語句分割器")
    st.write("上傳 PDF 並指定頁碼範圍，分句後可下載 Excel 檔。")

    pdf_file = st.file_uploader("請上傳 PDF 檔案", type="pdf")
    start_page = st.number_input("開始頁碼（從 1 起算）", min_value=1, step=1)
    end_page = st.number_input("結束頁碼（包含）", min_value=1, step=1)

    if pdf_file and end_page >= start_page:
        st.info("⏳ 正在處理 PDF，請稍候...")
        data = []
        with pdfplumber.open(pdf_file) as pdf:
            for i in range(start_page - 1, end_page):
                if i < len(pdf.pages):
                    page = pdf.pages[i]
                    text = page.extract_text()
                    sentences = split_sentences(text)
                    if sentences and re.match(r'^\d{1,3}/\d{1,3}$', sentences[0]):
                        sentences = sentences[1:]
                    for idx, s in enumerate(sentences, 1):
                        data.append({"頁碼": i+1, "語句編號": idx, "語句內容": s})

        df = pd.DataFrame(data)
        st.success("✅ 分句完成！預覽如下：")
        st.dataframe(df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button(
            label="📥 下載 Excel 檔",
            data=output.getvalue(),
            file_name="pdf_sentences.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == '__main__':
    main()
