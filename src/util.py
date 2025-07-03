from pathlib import Path
from textwrap import dedent
from typing import Iterator

import google.generativeai as genai


def dummy_ocr() -> str:
    """
    Returns:
        textbook (str): OCR結果のテキスト
    """

    import time

    time.sleep(5)

    return "aaaaaaaaaaa"


def dummy_gen_draft() -> Iterator[str]:
    """
    Returns:
        draft (Iterator[str]): レポートドラフトを1行ずつ返すイテレータ
    """
    import time

    time.sleep(5)

    f = ["avoubouvboaebovb\n"] * 100

    for line in f:
        time.sleep(0.2)
        yield line


def ocr(api_key: str, pdf_path: Path) -> str:
    """
    geminiによるocr

    Args:
        api_key (str): Google API key
        pdf_path (Path): PDFファイルのパス
    Returns:
        textbook (str): OCR結果のテキスト
    """
    # return dummy_ocr()

    genai.configure(api_key=api_key)

    model = genai.GenerativeModel("gemini-2.5-flash")
    prompt = "このPDFからすべてのテキストを抽出してください。数式はlatex形式で$$ x = 1 $$のように出力してください。出力はmarkdown形式でお願いします。"

    uploaded_file = genai.upload_file(path=str(pdf_path))
    file_part = {
        "file_data": {
            "mime_type": uploaded_file.mime_type,
            "file_uri": uploaded_file.uri,
        }
    }
    response = model.generate_content([file_part, prompt])

    return response.text


def gen_draft(api_key: str, textbook: str) -> Iterator[str]:
    """
    geminiによるレポートドラフト生成
    目的、原理、実験操作のみ

    Args:
        api_key (str): Google API key
        textbook (str): 実験テキストの内容（OCR結果）
    Returns:
        draft (Iterator[str]): レポートドラフト
    """
    # return dummy_gen_draft()
    genai.configure(api_key=api_key)

    prompt_instruction = dedent("""
    以下の実験テキストの内容をもとに、ルールを守ってレポートのドラフトを日本語で作成してください。
    レポートのドラフト以外の出力は絶対に避けてください。
    1. 目的, 2. 原理, 3. 実験の3つのセクションで構成してください。
    各セクションの内容は以下のようにしてください。
    - 目的: 実験テキストの目的を参考に、目的を簡潔に説明する。
    - 原理: 実験テキストの理論を参考に、原理を簡潔に説明する。必要に応じて数式を用いる。
    - 実験: 実験テキストの実験操作を参考に、実験の手順を簡潔に説明する。ただし、すべて過去形で書くこと。
    """)

    prompt_rules = dedent("""
    - 出力はマークダウン形式で書くこと。
    - 常態(～ある、である)で書き、箇条書きなどは避けること。
    - 数式は以下のような形でlatexのコードを示すこと。
    - 数式は必ず2つのドルマーク($$)で囲って示すこと。(例: $$ x_{A}$$ はモル分率) 

    (出力の例:

    このとき、モル中心の速度を基準とした成分Aのモル流束$$J_A$$は、Fickの第一法則により次式で与えられる。

    $$J_A = -C D_{AB} \frac{d x_A}{d z}$$

    ここで、$$C$$は全モル濃度[mol/m³]、$$D_{AB}$$は拡散係数[m²/s]、$$x_A$$は成分Aのモル分率[-]、$$z$$は位置[m]である。)
    """)

    # Geminiへのプロンプト
    prompt = dedent(f"""
    ## 指示
    {prompt_instruction}

    ## ルール
    {prompt_rules}

    ## 実験テキストの内容
    {textbook}
    """)

    model = genai.GenerativeModel("gemini-2.5-pro")

    # ストリーミング対応
    response = model.generate_content(
        prompt,
        stream=True,
    )

    for chunk in response:
        if hasattr(chunk, "text") and chunk.text:
            yield chunk.text
