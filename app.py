import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import base64

# --- ページ設定 ---
st.set_page_config(page_title="致知読書感想文アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("Step 1：記事読み込み（引用抽出） → Step 2：感想文作成")

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
# Step 1: 記事の読み込み
# ==========================================
st.header("Step 1. 記事画像のアップロード")
st.info("💡 複数の画像を一度に選んでアップロードしてください（PCならCtrlキーを押しながら選択）。")

uploaded_files = st.file_uploader("画像を選択（1ページ目、2ページ目...と複数可）", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files and st.button("🔍 記事の内容を詳しく抽出する", type="primary"):
    
    with st.spinner("AIが記事を読み、感想文に必要な箇所を抜き出しています..."):
        try:
            content_list = []
            
            # 【修正点】「スキャナー」ではなく「読書アシスタント」として振る舞わせる
            # これにより「読み取り拒否」を回避しつつ、正確な引用を引き出します
            system_prompt = """
            あなたは社内木鶏会のための読書アシスタントです。
            ユーザーが感想文を書くために、提供された記事画像の「詳細な内容」と「重要な文章」を抽出してください。

            【重要指示】
            1. 記事全体の流れを詳細に要約すること。
            2. 感想文の中で引用するために、著者の主張や印象的なエピソード部分は、勝手に要約せず「原文のまま」抜き出すこと。
            3. 「一字一句読めません」というエラーは出さず、読める範囲で最大限詳しくテキスト化すること。
            """
            
            content_list.append({"type": "text", "text": system_prompt})

            for f in uploaded_files:
                base64_image = encode_image(f)
                content_list.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}
                })

            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": content_list}],
                max_tokens=2500,
                temperature=0.2 # 少しだけ柔軟性を持たせて拒否を回避
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
    st.caption("感想文に使われる「素材」です。引用が間違っている場合はここで修正できます。")
    
    # ここで人間がチェック・修正できる
    edited_text = st.text_area("抽出テキスト（修正可）", st.session_state.extracted_text, height=400)
    st.session_state.extracted_text = edited_text

    # ==========================================
    # Step 2: 感想文の作成
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の作成")
    
    if st.button("✍️ 感想文を作成する"):
        with st.spinner("指定された条件で執筆中..."):
            try:
                writer_prompt = f"""
                あなたは税理士事務所の職員です。
                以下の【記事データ】を元に、社内木鶏会用の読書感想文を作成してください。

                【記事データ】
                {st.session_state.extracted_text}
                
                【作成条件】
                - 記事に書かれていないことを勝手に創作しないこと。
                - 上記データ内の「原文」を適切に引用しながら書くこと。
                - 構成：「①記事の要約」「②印象に残った言葉（引用）」「③自分の業務（税理士業務）への活かし方」
                - 文字数：{target_length}文字前後
                - 文体：「です・ます」調
                - Excelに貼り付けるため、段落ごとの改行のみとし、タイトルは不要。
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
            
            # 40文字分割書き込み
            lines = split_text(st.session_state.final_text, 40)
            
            start_row = 9
            # 既存のクリア
            for r in range(start_row, 50):
                ws[f"A{r}"].value = None
                ws[f"A{r}"].alignment = Alignment(wrap_text=False) # 一旦リセット

            for i, line in enumerate(lines):
                cell = ws[f"A{start_row + i}"]
                cell.value = line
                # 縮小して全体を表示
                cell.alignment = Alignment(shrink_to_fit=True, wrap_text=False)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            
            st.download_button(
                "📥 Excelをダウンロード", 
                out, 
                "致知感想文_完成.xlsx",
                type="primary"
            )
        except Exception as e:
            st.error(f"Excelエラー: {e}")
    else:
        st.warning("Excelフォーマットをアップロードしてください。")
