import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import base64

# --- ページ設定 ---
st.set_page_config(page_title="読書感想文作成アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("Step 1：内容確認 → Step 2：感想文作成（2段階方式）")

# --- APIキーの設定 ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("⚠️ OpenAI APIキーが設定されていません。Secretsを確認してください。")
    st.stop()

# --- 関数群 ---
def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

def split_text(text, chunk_size):
    return [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]

# --- セッション状態の管理 ---
if "extracted_text" not in st.session_state: st.session_state.extracted_text = ""
if "final_text" not in st.session_state: st.session_state.final_text = ""

# --- サイドバー設定 ---
with st.sidebar:
    st.header("⚙️ 設定")
    uploaded_template = st.file_uploader("感想文フォーマット(xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700], index=1)

# ==========================================
# Step 1: 記事の読み込みと正確な引用
# ==========================================
st.header("Step 1. 記事の読み込み")
st.info("💡 記事の画像（1ページ目、2ページ目...）をまとめて選択してアップロードしてください。")

# 複数ファイルを一括で受け取る設定
uploaded_files = st.file_uploader("画像を選択（複数可）", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files and st.button("🔍 記事を読み込んで内容を確認する", type="primary"):
    
    with st.spinner("AIが記事を一字一句、正確に読み取っています..."):
        try:
            content_list = []
            
            # 【重要】AIへの厳格な指示（勝手な解釈禁止）
            system_prompt = """
            あなたは高精度なOCR（文字認識）スキャナーです。
            提供された画像の文字を「一字一句正確に」読み取り、テキスト化してください。
            
            【厳守事項】
            1. 記事に書かれていないことを勝手に想像して追加しないこと。
            2. 記事の重要な文言は、省略せずにそのまま「引用」として抜き出すこと。
            3. 最後に、記事全体の要約を客観的な事実のみで作成すること。
            """
            
            content_list.append({"type": "text", "text": system_prompt})

            # アップロードされた全画像をリストに追加
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
                max_tokens=2000,
                temperature=0.0 # 0にすることで、創造性を排除し正確さを優先
            )
            
            # 結果を保存
            st.session_state.extracted_text = response.choices[0].message.content
            # 感想文はまだ作らないのでリセット
            st.session_state.final_text = "" 
            st.rerun()
            
        except Exception as e:
            st.error(f"読み取りエラー: {e}")

# ==========================================
# 読み取り結果の確認表示
# ==========================================
if st.session_state.extracted_text:
    st.markdown("---")
    st.subheader("📄 読み取り結果の確認")
    st.caption("AIが読み取った内容です。ここがおかしい場合は、画像を撮り直して再度Step 1を行ってください。")
    
    # ユーザーが修正できるようにテキストエリアにする
    edited_text = st.text_area("記事の内容（修正可能）", st.session_state.extracted_text, height=300)
    st.session_state.extracted_text = edited_text # 修正内容を保存

    # ==========================================
    # Step 2: 感想文の作成
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の作成")
    
    if st.button("✍️ この内容で感想文を作成する"):
        with st.spinner("感想文を執筆中..."):
            try:
                # 執筆用のプロンプト
                writer_prompt = f"""
                あなたは真面目な社員です。社内木鶏会で発表するための読書感想文を作成してください。
                
                【元となる記事の内容】
                {st.session_state.extracted_text}
                
                【作成条件】
                - 全体の文字数は{target_length}文字前後。
                - 上記の記事内容から正確に引用し、勝手な創作はしない。
                - 構成：
                  1. 記事の要約（簡潔に）
                  2. 記事を読んで特に心に残った言葉（正確に引用）
                  3. 自分の業務（税理士業務）や人生にどう生かすか
                - 文体は「です・ます」調。
                - タイトルは不要。
                - Excelのセルに入りきらないため、改行は最小限にする。
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
# Step 3: 完成とダウンロード
# ==========================================
if st.session_state.final_text:
    st.markdown("---")
    st.subheader("🎉 完成した感想文")
    st.text_area("完成プレビュー", st.session_state.final_text, height=300)

    if uploaded_template:
        try:
            wb = load_workbook(uploaded_template)
            ws = wb.active
            
            # 40文字区切り処理
            lines = split_text(st.session_state.final_text, 40)
            
            # A9セルから書き込み & 縮小設定
            start_row = 9
            # まず古い内容をクリア（念のためA9〜A30くらいまで）
            for r in range(start_row, 30):
                ws[f"A{r}"].value = None

            for i, line in enumerate(lines):
                cell = ws[f"A{start_row + i}"]
                cell.value = line
                cell.alignment = Alignment(shrink_to_fit=True, wrap_text=False)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            
            st.download_button(
                "📥 Excelファイルをダウンロード", 
                out, 
                "致知感想文.xlsx",
                type="primary"
            )
        except Exception as e:
            st.error(f"Excel保存エラー: {e}")
    else:
        st.warning("Excelフォーマットをアップロードすると、ここにダウンロードボタンが表示されます。")
