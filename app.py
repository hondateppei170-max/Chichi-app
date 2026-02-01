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
st.set_page_config(page_title="致知読書感想文アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("Step 1：画像解析 (Gemini 1.5 Flash) → Step 2：感想文執筆 (GPT-4o)")

# Excel書き込み設定
EXCEL_START_ROW = 9
CHARS_PER_LINE = 40

# ==========================================
# API設定
# ==========================================
try:
    # OpenAI
    openai_key = st.secrets.get("OPENAI_API_KEY")
    if not openai_key:
        st.error("⚠️ OpenAI APIキーが設定されていません。")
        st.stop()
    client = OpenAI(api_key=openai_key)

    # Google Gemini
    google_key = st.secrets.get("GOOGLE_API_KEY")
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
# サイドバー
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    uploaded_template = st.file_uploader("感想文フォーマット(.xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)

# ==========================================
# Step 1: 画像解析 (Gemini / 3記事対応)
# ==========================================
st.header("Step 1. 記事画像の解析 (Powered by Gemini)")
st.info("💡 複数の記事をタブごとに分けてアップロードしてください。Gemini 1.5 Flashで一括解析します。")

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

    if st.button("🔍 Geminiで全記事を解析する", type="primary"):
        with st.spinner("Geminiが画像を精読しています..."):
            try:
                # 入力リストの構築
                gemini_inputs = []
                
                # プロンプト
                system_prompt = """
                あなたはOCR（文字認識）のスペシャリストです。
                これから渡される雑誌『致知』の複数記事の画像から、テキスト情報を抽出してください。

                【抽出ルール】
                1. 記事ごとに「タイトル」「要約」「印象的な言葉（引用）」を抽出する。
                2. 引用文には必ず【掲載位置】を付記する（例：メイン記事 2枚目 右段）。
                3. 文字が読めない場合は「(判読不能)」と書く。ハルシネーション（嘘）は禁止。
                4. 以下の形式で出力すること。
                   ---
                   【記事1：メイン】
                   (内容)
                   【記事2】
                   (内容)
                   【記事3】
                   (内容)
                   ---
                """
                gemini_inputs.append(system_prompt)

                # 各タブの画像を処理
                article_labels = {"main": "【ここからメイン記事の画像】", "sub1": "【ここから記事2の画像】", "sub2": "【ここから記事3の画像】"}

                for key, files in files_dict.items():
                    if files:
                        # ファイル名順ソート
                        files.sort(key=lambda x: x.name)
                        
                        gemini_inputs.append(article_labels[key])
                        
                        for img_file in files:
                            # PIL Imageに変換
                            image = Image.open(img_file)
                            gemini_inputs.append(image)

                # Geminiモデル呼び出し
                # エラー回避のため 'gemini-1.5-flash-latest' を使用
                model = genai.GenerativeModel('gemini-1.5-flash-latest')
                
                response = model.generate_content(gemini_inputs)

                st.session_state.extracted_text = response.text
                st.session_state.final_text = ""
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
        "編集エリア（ここで修正した内容が感想文に使われます）", 
        st.session_state.extracted_text, 
        height=500
    )
    st.session_state.extracted_text = edited_text

    # ==========================================
    # Step 2: 感想文作成 (OpenAI使用)
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の執筆 (Powered by GPT-4o)")

    if st.button("✍️ 税理士事務所員として感想文を書く"):
        with st.spinner("GPT-4oが執筆中..."):
            try:
                writer_prompt = f"""
                あなたは税理士事務所の職員です。
                以下の【解析データ】を元に、社内木鶏会用の読書感想文を作成してください。

                【解析データ】
                {st.session_state.extracted_text}

                【構成】
                1. 記事の要約（複数の記事がある場合は、メインを中心にまとめる）
                2. 印象に残った言葉（解析データの引用元情報を活用し、正確に記載）
                3. 自分の業務（税理士補助・顧客対応・監査など）への具体的な活かし方

                【執筆条件】
                - 文字数：{target_length}文字前後
                - 文体：「です・ます」調
                - タイトル不要。段落ごとに改行を入れる。
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
