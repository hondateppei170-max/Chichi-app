import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
from PIL import Image

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="致知読書感想文アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ (Gemini 3.0版)")
st.caption("Step 1：画像解析 (Gemini 3 Flash) → Step 2：感想文執筆 (GPT-4o)")

# Excel書き込み設定
EXCEL_START_ROW = 9
CHARS_PER_LINE = 40

# ==========================================
# API設定
# ==========================================
try:
    # OpenAI (執筆用)
    openai_key = st.secrets.get("OPENAI_API_KEY")
    if not openai_key:
        st.warning("⚠️ OpenAI APIキーが設定されていません。")
    else:
        client = OpenAI(api_key=openai_key)

    # Google Gemini (画像解析用)
    google_key = st.secrets.get("GOOGLE_API_KEY")
    if not google_key:
        st.warning("⚠️ Google APIキーが設定されていません。")
    else:
        genai.configure(api_key=google_key)
    
except Exception as e:
    st.error(f"API設定エラー: {e}")
    st.stop()

# ==========================================
# 関数定義
# ==========================================
def split_text(text, chunk_size):
    """Excel用にテキストを指定文字数で分割"""
    if not text:
        return []
    clean_text = text.replace('\n', '　')
    return [clean_text[i:i+chunk_size] for i in range(0, len(clean_text), chunk_size)]

# ==========================================
# セッション状態
# ==========================================
if "extracted_text" not in st.session_state:
    st.session_state.extracted_text = ""
if "final_text" not in st.session_state:
    st.session_state.final_text = ""

# ==========================================
# サイドバー設定
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    uploaded_template = st.file_uploader("感想文フォーマット(.xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)
    
    st.markdown("---")
    st.caption("🔧 モデル設定")
    # デフォルト値を gemini-3-flash に設定
    model_id_input = st.text_input("GeminiモデルID", value="gemini-3-flash")
    st.caption("※Google AI Studio等で確認できるモデル名を入力")

# ==========================================
# Step 1: 画像解析 (Gemini 3 Flash)
# ==========================================
st.header("Step 1. 記事画像の解析")
st.info(f"💡 指定モデル「{model_id_input}」を使用して画像をOCR処理します。")

# 3つの記事に対応するタブ
tab1, tab2, tab3 = st.tabs(["📂 メイン記事", "📂 記事2 (任意)", "📂 記事3 (任意)"])

files_dict = {}

with tab1:
    files_dict["main"] = st.file_uploader("メイン記事の画像", type=['png', 'jpg', 'jpeg', 'webp'], accept_multiple_files=True, key="u1")
with tab2:
    files_dict["sub1"] = st.file_uploader("記事2の画像", type=['png', 'jpg', 'jpeg', 'webp'], accept_multiple_files=True, key="u2")
with tab3:
    files_dict["sub2"] = st.file_uploader("記事3の画像", type=['png', 'jpg', 'jpeg', 'webp'], accept_multiple_files=True, key="u3")

total_files = sum([len(f) for f in files_dict.values() if f])

if total_files > 0:
    st.write(f"📁 合計 {total_files}枚の画像を読み込みました")

    if st.button("🔍 画像解析を開始 (OCR)", type="primary"):
        with st.spinner(f"Gemini ({model_id_input}) が画像を読み込んでいます..."):
            try:
                gemini_inputs = []
                
                # システムプロンプト（OCR特化）
                system_prompt = """
あなたは、雑誌『致知』の紙面を完璧に読み取る高精度OCRエンジンです。
提供された全ての画像から、文字を一字一句漏らさず、ありのままに書き起こしてください。

【目的】
後続の処理でGPT-4oが記事を解析し、正確な「引用（掲載位置付き）」を作成するための元データを作成する。

【厳守ルール】
1. 完全な文字起こし（要約禁止）:
   - 要約や省略は一切禁止。書いてある文字を一字一句正確に書き起こすこと。
   - 縦書き（右上から左下）の文章の流れを正しく認識すること。

2. 位置情報のタグ付け（最重要）:
   - 後で「1枚目 右段」と特定できるように、テキストの前に位置情報を付記すること。
   - 画像ファイル名が判別できる場合は [ファイル名: xxx.jpg] と記載し、無理な場合は [画像N枚目] とする。

3. 記事ごとの区切り:
   - 提供される画像は複数の記事に分かれている場合があるため、入力データの区切り指示に従ってセクションを分けること。
"""
                gemini_inputs.append(system_prompt)

                # 画像データの準備
                for key, files in files_dict
