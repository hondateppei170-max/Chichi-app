import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import base64

# --- ページ設定 ---
st.set_page_config(page_title="読書感想文作成アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("社内木鶏会感想文 完全自動化ツール")

# --- APIキーの設定 (OpenAIのみ) ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("⚠️ OpenAI APIキーが設定されていません。Secretsを確認してください。")
    st.stop()

# --- 関数: 画像をBase64に変換 ---
def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

# --- 関数: 文章を指定文字数で分割する ---
def split_text(text, chunk_size):
    return [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]

# --- サイドバー設定 ---
with st.sidebar:
    st.header("⚙️ 出力設定")
    # 開始位置（A9）は固定コードにしていますが、変更可
    uploaded_template = st.file_uploader("感想文フォーマット(xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700], index=1)

# --- メイン画面 ---
st.info("Step 1: 雑誌の記事（画像）をアップロードしてください。")
uploaded_files = st.file_uploader("画像を選択", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files and st.button("🚀 感想文を作成する", type="primary"):
    
    # 1. GPT-4oによる読み取りと執筆
    with st.spinner("GPT-4oが記事を読み、感想文を書いています..."):
        try:
            content_list = []
            
            # プロンプト（指示書）
            system_prompt = f"""
            あなたは真面目な社員です。社内木鶏会で発表するための読書感想文を作成してください。
            
            【条件】
            - 文字数は{target_length}文字前後。
            - 記事の要約は短くまとめる。
            - 「①記事を読んで感じたこと」「②自分の業務や人生にどう生かすか」を必ず含める。
            - 文体は「です・ます」調。
            - タイトルや「感想文」という見出しは不要。本文のみ出力すること。
            """
            
            content_list.append({"type": "text", "text": system_prompt})

            # 画像の添付
            for f in uploaded_files:
                base64_image = encode_image(f)
                content_list.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}
                })

            # OpenAIへ送信
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": content_list}],
                max_tokens=1000,
                temperature=0.7
            )
            
            generated_text = response.choices[0].message.content
            st.session_state.final_text = generated_text
            st.success("✨ 完成しました！")
            
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

# --- Step 2: 確認とExcel出力 ---
if "final_text" in st.session_state and st.session_state.final_text:
    st.subheader("📝 作成された感想文")
    st.text_area("内容確認", st.session_state.final_text, height=300)

    if uploaded_template:
        try:
            # Excel処理
            wb = load_workbook(uploaded_template)
            ws = wb.active # 一番手前のシートを使います
            
            # 文章を40文字ごとに分割
            char_limit = 40
            lines = split_text(st.session_state.final_text, char_limit)
            
            # A9セルから順番に書き込む
            start_row = 9
            for i, line in enumerate(lines):
                target_cell = ws[f"A{start_row + i}"]
                target_cell.value = line
                
                # 「縮小して全体を表示」をONにする
                target_cell.alignment = Alignment(shrink_to_fit=True, wrap_text=False)

            # 保存用データ作成
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            
            st.download_button(
                label="📥 Excelに書き込んでダウンロード",
                data=out,
                file_name="致知感想文_完成.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Excel書き込みエラー: {e}")
            st.warning("Excelファイルが保護されていないか、形式が正しいか確認してください。")
    else:
        st.warning("👈 左のサイドバーで「感想文フォーマット(xlsx)」をアップロードすると、A9行から自動記入してダウンロードできます。")
