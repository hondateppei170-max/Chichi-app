import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from PIL import Image
import io

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="致知読書感想文アプリ(Gemini×GPT)", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("Step 1：画像解析 (Gemini 1.5 Flash) → Step 2：感想文執筆 (GPT-4o)")

# Excel書き込み設定
EXCEL_START_ROW = 9
CHARS_PER_LINE = 40

# ==========================================
# API設定 (Secretsから取得)
# ==========================================
try:
    # OpenAI設定
    openai_key = st.secrets.get("OPENAI_API_KEY")
    if not openai_key:
        st.error("⚠️ OpenAI APIキーが設定されていません。")
        st.stop()
    client = OpenAI(api_key=openai_key)

    # Google Gemini設定
    google_key = st.secrets.get("GOOGLE_API_KEY") # secrets.tomlに GOOGLE_API_KEY を設定してください
    if not google_key:
        st.error("⚠️ Google APIキーが設定されていません。")
        st.stop()
    genai.configure(api_key=google_key)
    
except Exception as e:
    st.error(f"API設定エラー: {e}")
    st.stop()

# ==========================================
# 関数定義
# ==========================================
def split_text(text, chunk_size):
    """Excel用にテキストを指定文字数で分割"""
    clean_text = text.replace('\n', '　') # 改行を全角スペースに置換
    return [clean_text[i:i+chunk_size] for i in range(0, len(clean_text), chunk_size)]

# ==========================================
# セッション状態
# ==========================================
if "extracted_text" not in st.session_state:
    st.session_state.extracted_text = ""
if "final_text" not in st.session_state:
    st.session_state.final_text = ""

# ==========================================
# サイドバー
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    uploaded_template = st.file_uploader("感想文フォーマット(.xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)

# ==========================================
# Step 1: 画像解析 (Gemini使用)
# ==========================================
st.header("Step 1. 記事画像の解析 (Powered by Gemini)")
st.info("💡 Gemini 1.5 Flashを使用し、大量の画像を一括高速解析します。")

uploaded_files = st.file_uploader(
    "画像をまとめて選択（ドラッグ＆ドロップ可）", 
    type=['png', 'jpg', 'jpeg', 'webp'], 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"📁 {len(uploaded_files)}枚の画像を読み込みました")

    if st.button("🔍 Geminiで画像を解析する", type="primary"):
        with st.spinner("Geminiが画像を読んでいます..."):
            try:
                # 1. ファイル名順にソート（重要）
                uploaded_files.sort(key=lambda x: x.name)

                # 2. 画像をPIL形式に変換してリスト化
                image_parts = []
                for file in uploaded_files:
                    image_parts.append(Image.open(file))

                # 3. Geminiへのプロンプト
                gemini_prompt = """
                あなたはOCRのスペシャリストです。
                添付された雑誌『致知』の全ページ画像を読み込み、以下の情報を抽出してください。

                【指示】
                1. 記事全体の詳細な要約を作成してください。
                2. 記事内の「重要な教え」や「印象的な言葉」を書き起こしてください。
                3. 書き起こしの際は、必ず「掲載位置」を付記してください（例：1枚目右段、3枚目写真キャプションなど）。
                4. 画像内の文字が読めない場合は無理に創作せず「(判読不能)」としてください。
                5. 嘘（ハルシネーション）は絶対禁止です。書いてあることだけを出力してください。
                """

                # 4. Geminiモデル呼び出し (gemini-1.5-flash は画像入力に強い)
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                # 画像とテキストをまとめて送信
                response = model.generate_content([gemini_prompt, *image_parts])

                st.session_state.extracted_text = response.text
                st.session_state.final_text = "" # リセット
                st.rerun()

            except Exception as e:
                st.error(f"Gemini解析エラー: {e}")

# ==========================================
# 解析結果の編集
# ==========================================
if st.session_state.extracted_text:
    st.markdown("---")
    st.subheader("📝 解析結果 (Gemini出力)")
    edited_text = st.text_area(
        "編集エリア（Step 2で使用されます）", 
        st.session_state.extracted_text, 
        height=500
    )
    st.session_state.extracted_text = edited_text

    # ==========================================
    # Step 2: 感想文作成 (OpenAI使用)
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の執筆 (Powered by GPT-4o)")

    if st.button("✍️ 感想文を作成する"):
        with st.spinner("GPT-4oが執筆中..."):
            try:
                writer_prompt = f"""
                あなたは税理士事務所の職員です。
                以下の【解析データ】を元に、社内木鶏会用の読書感想文を作成してください。

                【解析データ】
                {st.session_state.extracted_text}

                【構成】
                1. 記事の要約
                2. 印象に残った言葉（解析データの引用元情報を活用し、正確に記載）
                3. 自分の業務（税理士補助・顧客対応）への活かし方

                【条件】
                - 文字数：{target_length}文字前後
                - 文体：「です・ます」調
                - タイトル不要。段落ごとに改行。
                - 解析データにない内容は創作しないこと。
                """

                res = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": writer_prompt}],
                    temperature=0.7
                )

                st.session_state.final_text = res.choices[0].message.content
                st.rerun()

            except Exception as e:
                st.error(f"執筆エラー: {e}")

# ==========================================
# Step 3: Excel出力
# ==========================================
if st.session_state.final_text:
    st.markdown("---")
    st.subheader("🎉 完成＆Excel出力")
    st.text_area("完成テキスト", st.session_state.final_text, height=300)

    if uploaded_template:
        try:
            wb = load_workbook(uploaded_template)
            ws = wb.active

            # A9セル以降クリア
            for row in range(EXCEL_START_ROW, 100):
                ws[f"A{row}"].value = None

            # 40文字分割書き込み
            lines = split_text(st.session_state.final_text, CHARS_PER_LINE)
            
            for i, line in enumerate(lines):
                cell = ws[f"A{EXCEL_START_ROW + i}"]
                cell.value = line
                cell.alignment = Alignment(shrink_to_fit=True, wrap_text=False)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)

            st.download_button("📥 Excelダウンロード", out, "致知感想文.xlsx", type="primary")
        except Exception as e:
            st.error(f"Excel処理エラー: {e}")
    else:
        st.warning("Excelテンプレートをアップロードしてください。")
