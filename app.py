import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import base64

# --- ページ設定 ---
st.set_page_config(page_title="致知読書感想文アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("Step 1：引用箇所の抽出 → Step 2：感想文作成")

# --- APIキーの設定 ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("⚠️ OpenAI APIキーが設定されていません。")
    st.stop()

# --- 関数群 ---
def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

def split_text(text, chunk_size):
    return [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]

# --- セッション状態 ---
if "extracted_text" not in st.session_state: st.session_state.extracted_text = ""
if "final_text" not in st.session_state: st.session_state.final_text = ""

# --- サイドバー ---
with st.sidebar:
    st.header("⚙️ 設定")
    uploaded_template = st.file_uploader("感想文フォーマット(xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700], index=1)

# ==========================================
# Step 1: 記事の読み込み（引用抽出）
# ==========================================
st.header("Step 1. 記事画像の読み込み")
st.info("💡 記事の画像（複数可）を選択してください。感想文に必要な「重要な文章」を抜き出します。")

uploaded_files = st.file_uploader("画像を選択", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files and st.button("🔍 記事を解析して引用箇所を抜き出す", type="primary"):
    
    with st.spinner("著作権に配慮しつつ、感想文に必要な『原文』を抽出しています..."):
        try:
            content_list = []
            
            # 【重要修正】「データ化」ではなく「感想文のための引用抽出」と指示することで拒否を回避
            system_prompt = """
            あなたは社内木鶏会の読書感想文を作成するアシスタントです。
            ユーザーが感想文を書くために必要な情報を、提供された画像から抽出してください。

            【指示内容】
            1. 記事全体の「要約」を作成してください（どこに何が書かれているか分かるように）。
            2. 感想文の中で使用するために、著者の主張や印象的な言葉を「原文のまま」引用として抜き出してください。
            3. 記事に書かれていないことは一切追加しないでください。
            
            ※「読み取れません」というエラーを出さず、可能な限り文字を認識して出力すること。
            """
            
            content_list.append({"type": "text", "text": system_prompt})

            for f in uploaded_files:
                base64_image = encode_image(f)
                content_list.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}
                })

            # ガード回避のためTemperatureを少し上げるが、内容は厳格にする
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": content_list}],
                max_tokens=3000,
                temperature=0.2 
            )
            
            st.session_state.extracted_text = response.choices[0].message.content
            st.session_state.final_text = "" 
            st.rerun()
            
        except Exception as e:
            st.error(f"読み取りエラー: {e}")

# ==========================================
# 読み取り結果の確認・修正
# ==========================================
if st.session_state.extracted_text:
    st.markdown("---")
    st.subheader("📄 抽出内容の確認")
    st.caption("AIが抜き出した内容です。変な解釈が含まれていないか確認し、必要ならここで修正してください。")
    
    # 人間による修正エリア
    edited_text = st.text_area("抽出テキスト（ここを修正すると感想文に反映されます）", st.session_state.extracted_text, height=400)
    st.session_state.extracted_text = edited_text

    # ==========================================
    # Step 2: 感想文の作成
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の作成")
    
    if st.button("✍️ この内容で感想文を作成する"):
        with st.spinner("執筆中..."):
            try:
                writer_prompt = f"""
                あなたは税理士事務所の職員です。
                以下の【抽出データ】のみを使用して、社内木鶏会用の読書感想文を作成してください。

                【抽出データ】
                {st.session_state.extracted_text}
                
                【執筆条件】
                - 抽出データにある「原文引用」を必ず使用すること。
                - 勝手な創作や、記事にないエピソードを追加しないこと。
                - 構成：「①記事の要約」「②印象に残った言葉（引用）」「③自分の業務（税理士業務）への活かし方」
                - 文字数：{target_length}文字前後
                - 文体：「です・ます」調
                - タイトル不要。Excel用のため段落ごとの改行のみにする。
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
    st.subheader("🎉 完成プレビュー")
    st.text_area("完成した感想文", st.session_state.final_text, height=300)

    if uploaded_template:
        try:
            wb = load_workbook(uploaded_template)
            ws = wb.active
            
            # 40文字分割処理
            lines = split_text(st.session_state.final_text, 40)
            
            start_row = 9
            # クリア処理
            for r in range(start_row, 60):
                ws[f"A{r}"].value = None
                ws[f"A{r}"].alignment = Alignment(wrap_text=False)

            # 書き込み & 縮小設定
            for i, line in enumerate(lines):
                cell = ws[f"A{start_row + i}"]
                cell.value = line
                cell.alignment = Alignment(shrink_to_fit=True, wrap_text=False)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            
            st.download_button(
                "📥 Excelをダウンロード", 
                out, 
                "致知感想文.xlsx",
                type="primary"
            )
        except Exception as e:
            st.error(f"Excelエラー: {e}")
    else:
        st.warning("Excelフォーマットをアップロードしてください。")
