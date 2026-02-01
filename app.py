import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import base64

# --- ページ設定 ---
st.set_page_config(page_title="致知読書感想文アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("Step 1：記事の解析（引用元特定） → Step 2：感想文作成")

# --- APIキーの設定 ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("⚠️ OpenAI APIキーが設定されていません。Secretsを確認してください。")
    st.stop()

# --- 関数: 画像をBase64に変換 ---
def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

# --- 関数: 文章を指定文字数で分割 ---
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
# Step 1: 記事画像の読み込み（3枚まで個別対応）
# ==========================================
st.header("Step 1. 記事画像のアップロード")
st.info("💡 記事の画像を順番にアップロードしてください。")

col1, col2, col3 = st.columns(3)
with col1:
    img1 = st.file_uploader("記事 1枚目", type=['png', 'jpg', 'jpeg'], key="img1")
with col2:
    img2 = st.file_uploader("記事 2枚目", type=['png', 'jpg', 'jpeg'], key="img2")
with col3:
    img3 = st.file_uploader("記事 3枚目", type=['png', 'jpg', 'jpeg'], key="img3")

# アップロードされた画像をリストにまとめる
uploaded_images = []
if img1: uploaded_images.append(("1枚目", img1))
if img2: uploaded_images.append(("2枚目", img2))
if img3: uploaded_images.append(("3枚目", img3))

if uploaded_images and st.button("🔍 引用元を明記して解析する", type="primary"):
    
    with st.spinner("AIが画像の文字を読み、引用箇所と掲載位置を特定しています..."):
        try:
            content_list = []
            
            # 【重要】場所（ロケーション）を明記させる厳格な指示
            system_prompt = """
            あなたは「致知」の読書感想文を作成するための、厳格な記事解析アシスタントです。
            提供された画像から、感想文に必要な情報を抜き出してください。

            【絶対遵守の出力ルール】
            1. 記事全体の「詳細な要約」を作成すること。
            2. 「重要な文章」を抜き出す際は、必ず【掲載位置】を付記すること。
               例：「学ばざれば...」（1枚目 右段 5行目付近）
            3. 記事に書かれていないことは一切書かないこと（勝手な創作禁止）。
            4. 著者の名前や、記事内の人物名も正確に拾うこと。
            
            もし文字が不鮮明で読めない場合は、勝手に補完せず「（判読不能）」と書くこと。
            """
            
            content_list.append({"type": "text", "text": system_prompt})

            # 画像を順番にAIに見せる
            for label, img_file in uploaded_images:
                base64_image = encode_image(img_file)
                # 画像の前に「これは〇枚目の画像です」と注釈を入れる
                content_list.append({"type": "text", "text": f"【ここからは記事の『{label}』です】"})
                content_list.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}
                })

            # 解析実行（Temperature=0で事実のみ抽出）
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
# 読み取り結果の確認・修正
# ==========================================
if st.session_state.extracted_text:
    st.markdown("---")
    st.subheader("📄 解析結果（引用元の確認）")
    st.caption("記述に【1枚目 右段】などの場所が書かれているか確認してください。")
    
    edited_text = st.text_area("抽出テキスト（修正可）", st.session_state.extracted_text, height=500)
    st.session_state.extracted_text = edited_text

    # ==========================================
    # Step 2: 感想文の作成
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の作成")
    
    if st.button("✍️ 感想文を作成する"):
        with st.spinner("執筆中..."):
            try:
                writer_prompt = f"""
                あなたは税理士事務所の職員です。
                以下の【解析データ】を元に、社内木鶏会用の読書感想文を作成してください。

                【解析データ】
                {st.session_state.extracted_text}
                
                【執筆条件】
                - 解析データ内の「原文引用」を必ず使用し、記事に即した内容にすること。
                - 勝手な創作は禁止。記事にないエピソードは書かない。
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
            
            st.download_button("📥 Excelをダウンロード", out, "致知感想文.xlsx", type="primary")
        except Exception as e:
            st.error(f"Excelエラー: {e}")
    else:
        st.warning("Excelフォーマットをアップロードしてください。")
