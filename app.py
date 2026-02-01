import streamlit as st
import google.generativeai as genai
from openpyxl import load_workbook
import io

# --- ページ設定 ---
st.set_page_config(page_title="致知読書感想文作成アシスタント", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("鈴木尚剛税理士事務所 | 社内木鶏会感想文生成ツール")

# --- APIキーの設定 ---
try:
    # Streamlit CloudのSecretsからキーを取得
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    # モデルを「Flash」に設定（高速・安定・画像対応）
    model = genai.GenerativeModel('gemini-1.5-flash')
except Exception:
    st.error("⚠️ 設定エラー: APIキーが見つかりません。StreamlitのSettings > Secretsを確認してください。")
    st.stop()

# --- セッション状態の管理 ---
if "summary" not in st.session_state: st.session_state.summary = ""
if "final_text" not in st.session_state: st.session_state.final_text = ""

# --- サイドバー：設定 ---
with st.sidebar:
    st.header("⚙️ 設定")
    target_cell = st.text_input("開始セル", value="A9")
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700], index=1)
    uploaded_template = st.file_uploader("感想文フォーマット(xlsx)", type=["xlsx"])

# --- メイン機能 Step 1: 記事解析 ---
st.info("雑誌『致知』の記事（画像またはPDF）をアップロードしてください。")

# 画像とPDFの両方に対応
uploaded_files = st.file_uploader(
    "ファイルを選択（複数可）", 
    type=['png', 'jpg', 'jpeg', 'pdf'], 
    accept_multiple_files=True
)

if uploaded_files and st.button("記事を解析する", type="primary"):
    with st.spinner("Geminiが記事を読んでいます..."):
        try:
            # AIに渡すデータの準備
            prompt = "あなたはプロのライターです。提供された資料（雑誌記事）の「タイトル」と、300文字程度の「要約」を作成してください。"
            request_content = [prompt]
            
            for f in uploaded_files:
                # ファイルをバイトデータとして読み込む
                file_data = f.getvalue()
                # AIが読める形式に変換
                request_content.append({"mime_type": f.type, "data": file_data})
            
            # AIに送信
            response = model.generate_content(request_content)
            st.session_state.summary = response.text
            st.success("解析完了！")
            st.rerun()
            
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

# --- メイン機能 Step 2: 感想文生成 ---
if st.session_state.summary:
    st.subheader("📝 記事の要約")
    st.info(st.session_state.summary)
    
    st.divider()
    st.write("▼ 感想文の方向性を指示できます（例：『新人教育の難しさと絡めて』『感謝の心をテーマに』など）")
    user_instruction = st.text_input("指示（空欄のままでもOK）", key="instruction")

    if st.button("✨ この内容で感想文を作成する"):
        with st.spinner("感想文を執筆中..."):
            try:
                final_prompt = f"""
                あなたは税理士事務所の職員です。以下の要約と指示を元に、社内木鶏会で発表するための読書感想文を作成してください。
                
                【記事要約】: {st.session_state.summary}
                【ユーザーの指示】: {user_instruction}
                
                【条件】: 
                1. 文字数は {target_length} 文字前後。
                2. 「①感じたこと」「②人生・仕事（税理士業務）にどう生かすか」の要素を含める。
                3. 文体は「です・ます」調で、真摯なトーンで。
                4. タイトルは不要。本文のみ出力。
                """
                
                res = model.generate_content(final_prompt)
                st.session_state.final_text = res.text
                st.rerun()
            except Exception as e:
                st.error(f"作成中にエラーが発生しました: {e}")

# --- メイン機能 Step 3: 出力 ---
if st.session_state.final_text:
    st.subheader("🎉 完成した感想文")
    st.text_area("内容確認", st.session_state.final_text, height=300)
    
    if uploaded_template:
        try:
            wb = load_workbook(uploaded_template)
            ws = wb.active
            ws[target_cell] = st.session_state.final_text
            
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            
            st.download_button(
                label="📥 Excelファイルをダウンロード",
                data=out,
                file_name="致知感想文.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error("Excelファイルへの書き込みに失敗しました。フォーマットを確認してください。")
    else:
        st.warning("⚠️ サイドバーからExcelフォーマットをアップロードすると、直接ファイルに書き込めます。")
