import streamlit as st
import subprocess
import sys
import io

# --- 🛠️ 強制修復エリア（ここが重要です） ---
# システムが古い道具を使わないよう、アプリ起動時に強制的に最新版を入れます
try:
    import google.generativeai
    # バージョンが古い、または入っていない場合はエラーを起こして修復に進む
    if google.generativeai.__version__ < "0.8.3":
        raise ImportError
except ImportError:
    # 画面にメッセージを出してインストール開始
    st.write("🔧 AIの準備をしています...（初回のみ時間がかかります）")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "google-generativeai>=0.8.3", "openpyxl"])
    st.rerun() # インストール後に再起動

# ---------------------------------------------

import google.generativeai as genai
from openpyxl import load_workbook

# --- ページ設定 ---
st.set_page_config(page_title="致知読書感想文作成アシスタント", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("鈴木尚剛税理士事務所 | 社内木鶏会感想文生成ツール")

# --- APIキーの設定 ---
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    # 100%確実に動くモデルを指定
    model = genai.GenerativeModel('gemini-1.5-flash')
except Exception as e:
    st.error(f"設定エラー: APIキーが見つかりません。Settings > Secretsを確認してください。\n詳細: {e}")
    st.stop()

# --- セッション状態 ---
if "summary" not in st.session_state: st.session_state.summary = ""
if "final_text" not in st.session_state: st.session_state.final_text = ""

# --- サイドバー ---
with st.sidebar:
    st.header("⚙️ 設定")
    target_cell = st.text_input("開始セル", value="A9")
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700], index=1)
    uploaded_template = st.file_uploader("感想文フォーマット(xlsx)", type=["xlsx"])

# --- メイン機能 Step 1: 記事解析 ---
st.info("雑誌『致知』の記事（画像またはPDF）をアップロードしてください。")

uploaded_files = st.file_uploader(
    "ファイルを選択（複数可）", 
    type=['png', 'jpg', 'jpeg', 'pdf'], 
    accept_multiple_files=True
)

if uploaded_files and st.button("記事を解析する", type="primary"):
    with st.spinner("Geminiが記事を読んでいます..."):
        try:
            prompt = "あなたはプロのライターです。提供された資料（雑誌記事）の「タイトル」と、300文字程度の「要約」を作成してください。"
            request_content = [prompt]
            
            for f in uploaded_files:
                request_content.append({"mime_type": f.type, "data": f.getvalue()})
            
            response = model.generate_content(request_content)
            st.session_state.summary = response.text
            st.success("解析完了！")
            st.rerun()
            
        except Exception as e:
            st.error(f"解析エラーが発生しました: {e}")
            st.info("【ヒント】Google APIキーが正しいか、もう一度確認してください。")

# --- メイン機能 Step 2: 感想文生成 ---
if st.session_state.summary:
    st.subheader("📝 記事の要約")
    st.info(st.session_state.summary)
    
    st.divider()
    user_instruction = st.text_input("感想文の方向性（例：『感謝の心をテーマに』など、空欄でもOK）", key="instruction")

    if st.button("✨ 感想文を作成する"):
        with st.spinner("感想文を執筆中..."):
            try:
                final_prompt = f"""
                以下の要約と指示を元に、社内木鶏会で発表するための読書感想文を作成してください。
                【記事要約】: {st.session_state.summary}
                【ユーザーの指示】: {user_instruction}
                【条件】: 
                - 文字数は {target_length} 文字前後。
                - 「①感じたこと」「②人生・仕事（税理士業務）にどう生かすか」を含める。
                - 文体は「です・ます」調。タイトル不要。
                """
                res = model.generate_content(final_prompt)
                st.session_state.final_text = res.text
                st.rerun()
            except Exception as e:
                st.error(f"作成エラー: {e}")

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
            st.download_button("📥 Excelダウンロード", out.getvalue(), "致知感想文.xlsx")
        except Exception as e:
            st.error(f"Excel保存エラー: {e}")
    else:
        st.warning("Excelフォーマットをアップロードすると、直接ファイルに書き込めます。")
