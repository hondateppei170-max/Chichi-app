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
    # 【修正】デフォルト値を gemini-3-flash-preview に変更
    model_id_input = st.text_input("GeminiモデルID", value="gemini-3-flash-preview")
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
                
                # ==========================================
                # プロンプト：上下2段組み対応・読み順指定
                # ==========================================
                system_prompt_text = (
                    "あなたは、雑誌『致知』の紙面を完璧に読み取る高精度OCRエンジンです。\n"
                    "提供された画像は、基本的に「上下2段組み」のレイアウトになっています。\n"
                    "以下の読み取り順序を厳守して書き起こしてください。\n\n"
                    "【読み取り順序の絶対ルール】\n"
                    "1. ページ構成の認識:\n"
                    "   - 画像を「上半分（上段）」と「下半分（下段）」に分けて認識すること。\n"
                    "2. 読み進める順番:\n"
                    "   - まず、【上段】の文章を「右から左へ」全て読み取る。\n"
                    "   - 次に、【下段】の文章を「右から左へ」全て読み取る。\n"
                    "   - ※絶対に左側の段を上から下まで一気に読んではいけない（上下が混ざらないようにする）。\n\n"
                    "【出力形式】\n"
                    "各画像のテキストの前に、必ず以下の位置タグを付けること。\n"
                    "   [画像N枚目]\n"
                    "   <上段> ...ここに上段のテキスト...\n"
                    "   <下段> ...ここに下段のテキスト...\n\n"
                    "【その他のルール】\n"
                    "   - 要約や省略は一切禁止。一字一句ありのままに書き起こす。\n"
                    "   - 縦書き特有の「右上から始まり左下へ終わる」流れを守る。\n"
                    "   - 画像ファイル名がわかる場合は [ファイル名: xxx.jpg] と記載する。"
                )
                gemini
