import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from openpyxl import load_workbook
import io

# --- ページ設定 ---
st.set_page_config(page_title="致知読書感想文マスター", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ (Gemini × ChatGPT)")
st.caption("鈴木尚剛税理士事務所 | 完全自動化ツール")

# --- APIキーの設定 ---
# 1. Gemini (読み取り担当)
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    model_gemini = genai.GenerativeModel('gemini-1.5-flash')
except Exception:
    st.error("⚠️ Google APIキーの設定が必要です。")

# 2. ChatGPT (執筆担当)
try:
    client_gpt = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("⚠️ OpenAI APIキーの設定が必要です。Secretsを確認してください。")

# --- セッション状態 ---
if "extracted_text" not in st.session_state: st.session_state.extracted_text = ""
if "final_text" not in st.session_state: st.session_state.final_text = ""

# --- サイドバー設定 ---
with st.sidebar:
    st.header("⚙️ 出力設定")
    target_cell = st.text_input("Excelの開始セル", value="A9")
    target_length = st.selectbox("文字数", [300, 400, 500, 600, 700], index=1)
    uploaded_template = st.file_uploader("感想文フォーマット(xlsx)", type=["xlsx"])

# --- メイン機能 ---
st.info("雑誌『致知』の記事（画像またはPDF）をアップロードしてください。")
uploaded_files = st.file_uploader("ファイルを選択", type=['png', 'jpg', 'jpeg', 'pdf'], accept_multiple_files=True)

if uploaded_files and st.button("🚀 自動作成スタート", type="primary"):
    
    # Step 1: Geminiで文字を読む
    with st.spinner("👀 Geminiが記事を読んでいます..."):
        try:
            prompt = "この資料の文字をすべて読み取って、内容を詳細にテキスト化してください。"
            request_content = [prompt]
            for f in uploaded_files:
                request_content.append({"mime_type": f.type, "data": f.getvalue()})
            
            response_gemini = model_gemini.generate_content(request_content)
            st.session_state.extracted_text = response_gemini.text
        except Exception as e:
            st.error(f"読み取りエラー: {e}")
            st.stop()

    # Step 2: ChatGPTで感想文を書く
    with st.spinner("✍️ ChatGPTが感想文を執筆中..."):
        try:
            system_prompt = "あなたは税理士事務所の真面目な職員です。社内木鶏会で発表するための読書感想文を作成してください。"
            user_prompt = f"""
            以下の記事内容を元に、読書感想文を書いてください。

            【記事の内容】:
            {st.session_state.extracted_text}

            【条件】:
            - 文字数は {target_length} 文字前後。
            - 「①記事を読んで感じたこと」「②自分の業務（税理士業務）や人生にどう生かすか」を含める。
            - 文体は「です・ます」調。タイトルは不要。
            """

            response_gpt = client_gpt.chat.completions.create(
                model="gpt-4o", # 最新の高精度モデル
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.7,
            )
            st.session_state.final_text = response_gpt.choices[0].message.content
            st.success("✨ 完成しました！")
            st.rerun()
            
        except Exception as e:
            st.error(f"執筆エラー: {e}")

# --- 結果表示とダウンロード ---
if st.session_state.final_text:
    st.subheader("🎉 完成した感想文")
    st.text_area("内容確認", st.session_state.final_text, height=400)
    
    # Excelダウンロードボタン
    if uploaded_template:
        try:
            wb = load_workbook(uploaded_template)
            ws = wb.active
            ws[target_cell] = st.session_state.final_text
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            st.download_button("📥 Excelファイルでダウンロード", out, "致知感想文.xlsx")
        except Exception as e:
            st.error("Excel書き込みエラー")
    else:
        st.warning("⚠️ サイドバーでExcelフォーマットをアップロードすると、直接ファイルに書き込めます。")
