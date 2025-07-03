from pathlib import Path
import tempfile

import streamlit as st

from util import ocr, gen_draft, dummy_gen_draft


def main():
    st.title("Report Draft Generator β version")
    st.write("This app generates a draft report based on the provided PDF textbook.")

    # Input for API key
    api_key = st.text_input("Enter your Gemini API key:", type="password")

    # PDFファイル限定のアップローダー
    uploaded_file = st.file_uploader("Upload your textbook PDF file", type=["pdf"])

    if st.button("Generate Draft"):
        if api_key and uploaded_file:
            # 進捗表示用のプレースホルダー
            progress = st.empty()
            with st.spinner("処理中..."):
                progress.info("OCR（テキスト抽出）を実行中です...")
                with tempfile.NamedTemporaryFile(
                    delete=True, suffix=".pdf"
                ) as tmp_file:
                    tmp_file.write(uploaded_file.getbuffer())
                    tmp_file.flush()
                    pdf_path = Path(tmp_file.name)

                    # ocr
                    textbook = ocr(api_key, pdf_path)

                progress.info("レポートドラフトを生成中です...")
                st.write("### Generated Draft")
                placeholder = st.empty()
                draft_text = ""

                for chunk in gen_draft(api_key, textbook):
                    draft_text += chunk
                    placeholder.markdown(draft_text)
                # for chunk in dummy_gen_draft():
                #     draft_text += chunk
                #     placeholder.markdown(draft_text)

                # コピーボタン
                st.components.v1.html(
                    f"""
                    <button onclick="navigator.clipboard.writeText(`{draft_text}`)">レポートのドラフトをコピー</button>
                    """,
                    height=50,
                )
                st.components.v1.html(
                    f"""
                    <button onclick="navigator.clipboard.writeText(`{textbook}`)">実験テキストをコピー</button>
                    """,
                    height=50,
                )
            progress.empty()
        else:
            st.error("Please provide both API key and a PDF file.")


if __name__ == "__main__":
    main()
