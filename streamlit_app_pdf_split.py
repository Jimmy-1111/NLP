import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# ---------- 分句函式 ---------- #
def split_sentences(text: str):
    """將整頁文字分句，並排除短句 / 純符號 / 關鍵字"""
    if not text:
        return []

    # 1) 先把所有非空行接在一起，減少硬換行的碎裂
    cleaned = " ".join(line.strip() for line in text.splitlines() if line.strip())

    # 2) 用中英文句號、驚嘆號、問號斷句
    raw_sentences = re.split(r'(?<=[。．！？.!?])', cleaned)

    # 3) 不想要的關鍵字
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

        # 3-a) 過短或純數字/符號句子直接跳過
        if len(s) < 5 or re.fullmatch(r'^[\d\W\s]+$', s):
            continue

        # 3-b) 排除關鍵字
        if any((kw in s) if isinstance(kw, str) else kw.search(s) for kw in exclude_keywords):
            continue

        sentences.append(s)
    return sentences


# ---------- Streamlit 主程式 ---------- #
def main():
    st.title("📄 PDF 語句分割器（精簡修補版）")
    st.write("上傳 PDF → 選擇頁碼範圍 → 分句 → 下載 Excel")

    # 初始化
    if "ranges" not in st.session_state:
        st.session_state["ranges"] = []

    # 上傳 PDF
    pdf_file = st.file_uploader("請上傳 PDF 檔案", type="pdf")

    # 頁碼範圍
    col1, col2 = st.columns(2)
    with col1:
        start_page = st.number_input("開始頁碼 (1 起算)", min_value=1, step=1, value=1)
    with col2:
        end_page = st.number_input("結束頁碼 (含)", min_value=1, step=1, value=1)

    if st.button("➕ 新增範圍"):
        if end_page >= start_page:
            st.session_state["ranges"].append((start_page, end_page))
        else:
            st.warning("⚠️ 結束頁碼不可小於開始頁碼")

    if st.session_state["ranges"]:
        st.markdown("🗂 **已選範圍**")
        for i, (s, e) in enumerate(st.session_state["ranges"], 1):
            st.write(f"{i}. 第 {s}–{e} 頁")

    # 命名資訊
    company = st.text_input("企業名稱")
    year    = st.text_input("年份 (4 位)")
    month   = st.text_input("月份 (可空白)")
    day     = st.text_input("日期 (可空白)")

    filename = "_".join(filter(None, [company.strip(), year.strip(), month.strip(), day.strip()])) or "output"
    filename += ".xlsx"

    # --------------- 開始處理 --------------- #
    if st.button("🚀 開始分句") and pdf_file and st.session_state["ranges"]:
        st.info("⏳ 解析中，請稍候…")

        data = []
        with pdfplumber.open(pdf_file) as pdf:
            for (s_page, e_page) in st.session_state["ranges"]:
                for i in range(s_page - 1, e_page):
                    if i >= len(pdf.pages):
                        continue
                    page_text = pdf.pages[i].extract_text()
                    for idx, sent in enumerate(split_sentences(page_text), 1):
                        data.append({"頁碼": i + 1, "語句編號": idx, "語句內容": sent})

        df = pd.DataFrame(data)
        if df.empty:
            st.error("😕 沒抓到任何句子，請檢查 PDF 是否為掃描影像或頁碼範圍是否正確")
            return

        st.success("✅ 完成！以下為預覽")
        st.dataframe(df, use_container_width=True)

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 下載 Excel", data=out.getvalue(),
                           file_name=filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # 重置範圍
        st.session_state["ranges"] = []


if __name__ == "__main__":
    main()
