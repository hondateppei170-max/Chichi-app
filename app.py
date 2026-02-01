import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import base64

# --- ページ設定 ---
st.set_page_config(page_title="致知読書感想文アプリ（厳格版）", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("Step 1：正確な読み取り確認 → Step 2：感想文作成")

# --- APIキーの設定 ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("⚠️ OpenAI APIキーの設定が必要です。Secretsを確認してください。")
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
# Step 1: 記事の厳格な読み取り
# ==========================================
st.header("Step 1. 記事画像の読み込み（根拠の抽出）")
st.info("💡 記事の画像を選択してください（複数可）。AIが「書いてあることだけ」を抜き出します。")

uploaded_files = st.file_uploader("画像を選択", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files and st.button("🔍 記事を解析する（解釈禁止モード）", type="primary"):
    
    with st.spinner("AIが主観を排除して記事を読み取っています..."):
        try:
            content_list = []
            
            # 【重要】AIへの厳格な指示（Temperature=0で運用）
            system_prompt = """
            あなたは「書かれている文字を正確にデータ化する」厳格なアシスタントです。
            提供された雑誌記事の画像から、感想文に必要な情報を抜き出してください。

            【絶対厳守のルール】
            1. 「要約」を作成する際は、必ずその根拠となる文章が画像のどこにあるか（例：1枚目右段、2枚目左段など）を明記すること。
            2. 著者の主張や名言を抜き出す際は、一言一句変えず、勝手な要約をせずに「原文のまま」引用すること。
            3. 記事に書かれていない情報（一般的な知識やネットの情報）は一切混ぜないこと。
            4. 読めない文字がある場合は、勝手に補完せず「（判読不能）」と書くこと。
            """
            
            content_list.append({"type": "text", "text": system_prompt})

            for f in uploaded_files:
                base64_image = encode_image(f)
                content_list.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}
                })

            # Temperatureを0に設定＝「創造性ゼロ・事実のみ」
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": content_list}],
                max_tokens=3000,
                temperature=0.0
            )
            
            st.session_state.extracted_text = response.choices[0].message.content
            st.session_state.final_text = "" 
            st.rerun()
            
        except Exception as e:
            st.error(f"読み取りエラー: {e}")

# ==========================================
# 読み取り結果の確認・修正（ここが重要）
# ==========================================
if st.session_state.extracted_text:
    st.markdown("---")
    st.subheader("📄 読み取り結果の確認")
    st.warning("⚠️ 以下の内容に「勝手な解釈」が含まれていないか確認してください。修正も可能です。")
    
    # ユーザーが修正できるエリア
    edited_text = st.text_area("抽出されたテキスト", st.session_state.extracted_text, height=500)
    st.session_state.extracted_text = edited_text

    # ==========================================
    # Step 2: 感想文の作成
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の作成")
    
    if st.button("✍️ 上記の「事実」のみに基づいて感想文を作成する"):
        with st.spinner("執筆中..."):
            try:
                writer_prompt = f"""
                あなたは税理士事務所の職員です。
                以下の【確定した記事データ】のみを使用して、社内木鶏会用の読書感想文を作成してください。

                【確定した記事データ】
                {st.session_state.extracted_text}
                
                【執筆条件】
                - 記事データにない情報は一切書かないこと（勝手な補足禁止）。
                - 記事内の言葉を引用する場合は、一言一句正確に引用すること。
                - 構成：
                  1. 記事の要約（短く）
                  2. 特に感銘を受けた言葉（原文引用）
                  3. それを税理士業務や自分の人生にどう活かすか（ここだけは自分の決意を書く）
                - 文字数：{target_length}文字前後
                - 文体：「です・ます」調
                - タイトル不要。
                """

                res = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": writer_prompt}],
                    temperature=0.7 # 文章の自然さのために少しだけ上げるが、ソースは厳守させる
                )
                
                st.session_state.final_text = res.choices[0].message.content
                st.rerun()
                
            except Exception as e:
                st.error(f"執筆エラー: {e}")

# ==========================================
# Step 3: Excel出力（40文字分割＆縮小）
# ==========================================
if st.session_state.final_text:
    st.markdown("---")
    st.subheader("🎉 完成プレビュー")
    st.text_area("完成した感想文", st.session_state.final_text, height=300)

    if uploaded_template:
        try:
            wb = load_workbook(uploaded_template)
            ws = wb.active
            
            # 40文字区切り処理
            lines = split_text(st.session_state.final_text, 40)
            
            start_row = 9
            # 書き込み前に古いデータをクリア
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
            st.error(f"Excel保存エラー: {e}")
    else:
        st.warning("Excelフォーマットをアップロードしてください。")
