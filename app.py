import streamlit as st
import google.generativeai as genai
from openpyxl import load_workbook
import io
import datetime

# ページ設定
st.set_page_config(page_title="致知読書感想文作成アシスタント", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("鈴木尚剛税理士事務所 | 社内木鶏会感想文生成ツール")

# APIキーの取得（Streamlit Cloudの金庫から読み込む安全な方法）
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    model = genai.GenerativeModel('gemini-1.5-flash')
except Exception:
    st.error("設定エラー: APIキーが見つかりません。")
    st.stop()

# セッション管理
if "chat_history" not in st.session_state: st.session_state.chat_history = []
if "summary" not in st.session_state: st.session_state.summary = ""
if "final_text" not in st.session_state: st.session_state.final_text = ""

# サイドバー設定
with st.sidebar:
    st.header("設定")
    target_cell = st.text_input("開始セル", value="A9")
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700], index=1)
    uploaded_template = st.file_uploader("感想文フォーマット(xlsx)", type=["xlsx"])

# メイン機能：記事読み込み
st.info("記事の画像をアップロードしてください。")
uploaded_imgs = st.file_uploader("画像を選択（複数可）", accept_multiple_files=True)

if uploaded_imgs and st.button("記事を解析"):
    with st.spinner("Geminiが記事を読んでいます..."):
        prompt = "この画像のタイトルと、300文字程度の要約を作成してください。"
        request_content = [prompt]
        for f in uploaded_imgs:
            request_content.append({"mime_type": f.type, "data": f.getvalue()})
        
        response = model.generate_content(request_content)
        st.session_state.summary = response.text
        st.rerun()

# 結果表示と感想文作成
if st.session_state.summary:
    st.subheader("記事の要約")
    st.write(st.session_state.summary)
    
    # チャット機能（シンプル化）
    st.divider()
    st.write("感想の方向性を指示できます（例：「新人教育の悩みを絡めて」など）")
    user_instruction = st.text_input("指示（空欄でもOK）", key="instruction")

    if st.button("感想文を作成する", type="primary"):
        with st.spinner("執筆中..."):
            final_prompt = f"""
            以下の情報を元に、社内木鶏会の感想文を書いてください。
            【要約】: {st.session_state.summary}
            【ユーザーの指示】: {user_instruction}
            【条件】: 文字数は{target_length}文字前後。仕事（税理士業務）への情熱を含める。
            """
            res = model.generate_content(final_prompt)
            st.session_state.final_text = res.text

# 出力エリア
if st.session_state.final_text:
    st.text_area("完成した感想文", st.session_state.final_text, height=300)
    
    if uploaded_template:
        wb = load_workbook(uploaded_template)
        ws = wb.active
        ws[target_cell] = st.session_state.final_text
        out = io.BytesIO()
        wb.save(out)
        st.download_button("Excelダウンロード", out.getvalue(), "感想文.xlsx")
    else:
        st.warning("左側のサイドバーからExcelをアップロードすると、直接書き込めます。")
