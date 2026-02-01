import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import base64

# ==========================================
# ページ設定・定数定義
# ==========================================
st.set_page_config(page_title="致知読書感想文アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ v2")
st.caption("Step 1：記事の解析（事実抽出） → Step 2：感想文執筆（税理士事務所向け）")

# Excel書き込み設定
EXCEL_START_ROW = 9
CHARS_PER_LINE = 40

# ==========================================
# 関数定義
# ==========================================

def get_openai_client():
    api_key = st.secrets.get("OPENAI_API_KEY")
    if not api_key:
        st.error("⚠️ OpenAI APIキーが設定されていません。")
        st.stop()
    return OpenAI(api_key=api_key)

def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

def split_text(text, chunk_size):
    clean_text = text.replace('\n', '　')
    return [clean_text[i:i+chunk_size] for i in range(0, len(clean_text), chunk_size)]

# ==========================================
# セッション状態
# ==========================================
if "extracted_text" not in st.session_state:
    st.session_state.extracted_text = ""
if "final_text" not in st.session_state:
    st.session_state.final_text = ""

client = get_openai_client()

# ==========================================
# サイドバー設定
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    uploaded_template = st.file_uploader("感想文フォーマット(.xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)

# ==========================================
# Step 1: 複数記事画像の読み込みと解析
# ==========================================
st.header("Step 1. 記事画像の解析")
st.warning("⚠️ 画像が不鮮明だとAIが内容を勝手に創作する場合があります。明るく鮮明な画像を使用してください。")

# タブで記事ごとにアップロード欄を分ける
tab1, tab2, tab3 = st.tabs(["📂 メイン記事", "📂 記事2 (任意)", "📂 記事3 (任意)"])

files_dict = {}

with tab1:
    files_dict["main"] = st.file_uploader("メイン記事の画像 (複数可)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="u1")
with tab2:
    files_dict["sub1"] = st.file_uploader("2つ目の記事の画像", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="u2")
with tab3:
    files_dict["sub2"] = st.file_uploader("3つ目の記事の画像", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="u3")

# 全てのアップロードファイルを確認
total_files = sum([len(f) for f in files_dict.values() if f])

if total_files > 0:
    if st.button("🔍 画像を解析する（創作禁止モード）", type="primary"):
        with st.spinner("AIが画像を精読しています...（時間がかかります）"):
            try:
                content_list = []
                
                # システムプロンプト：事実のみを抽出するよう厳格化
                system_prompt = """
                あなたはOCR（光学文字認識）の専門家です。
                提供された雑誌『致知』の画像から、文字情報を正確に読み取ってください。

                【最重要禁止事項】
                - 記事に書かれていない内容（一般的な知識や推測）を絶対に追記してはならない。
                - 画像が不鮮明で読めない場合は、勝手に補完せず「（判読不能）」と記述すること。
                - ハルシネーション（嘘の記述）は厳禁です。

                【出力形式】
                各記事ごとに以下の形式で出力してください。
                1. 記事タイトル（見える範囲で）
                2. 登場人物名（正確に）
                3. 詳細な要約（記事にある事実のみで構成）
                4. 重要な引用文（掲載位置を付記：例「〜である」（2枚目 右段））
                """
                content_list.append({"type": "text", "text": system_prompt})

                # 各タブの画像を順番に追加
                article_labels = {"main": "【メイン記事】", "sub1": "【2つ目の記事】", "sub2": "【3つ目の記事】"}
                
                for key, files in files_dict.items():
                    if files:
                        # ファイル名順にソート
                        files.sort(key=lambda x: x.name)
                        content_list.append({"type": "text", "text": f"\n\n=== ここから{article_labels[key]} ===\n"})
                        
                        for i, img_file in enumerate(files):
                            base64_img = encode_image(img_file)
                            content_list.append({
                                "type": "text", 
                                "text": f"\n[{article_labels[key]} {i+1}枚目 (ファイル名: {img_file.name})]\n"
                            })
                            content_list.append({
                                "type": "image_url",
                                "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}
                            })

                # 解析実行
                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": content_list}],
                    max_tokens=4000,
                    temperature=0.0  # 厳密に事実のみ
                )

                st.session_state.extracted_text = response.choices[0].message.content
                st.session_state.final_text = "" 
                st.rerun()

            except Exception as e:
                st.error(f"解析エラー: {e}")

# ==========================================
# 解析結果の確認・修正
# ==========================================
if st.session_state.extracted_text:
    st.markdown("---")
    st.subheader("📝 解析結果の確認")
    st.warning("内容が記事と合っているか必ず確認してください。違っている場合はここで修正してください。")
    
    edited_text = st.text_area(
        "解析テキスト（修正用）", 
        st.session_state.extracted_text, 
        height=500
    )
    st.session_state.extracted_text = edited_text

    # ==========================================
    # Step 2: 感想文の執筆
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の執筆")

    if st.button("✍️ 感想文を作成する"):
        with st.spinner("執筆中..."):
            try:
                writer_prompt = f"""
                あなたは税理士事務所の職員です。
                以下の【解析データ】のみを使用して、社内木鶏会用の読書感想文を作成してください。

                【解析データ】
                {st.session_state.extracted_text}

                【厳守ルール】
                - 解析データに含まれていない情報は一切書かないこと（嘘を混ぜない）。
                - 複数の記事がある場合は、それらを関連付けてまとめるか、メイン記事を中心に構成する。
                - 構成：「①要約」「②印象に残った言葉（引用）」「③業務（税理士業務）への活かし方」。
                - 文字数：{target_length}文字前後。
                - 文体：「です・ます」調。
                - タイトル不要、段落ごとに改行。
                """

                res = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": writer_prompt}],
                    temperature=0.5 # 執筆時は少し自然にするが、創作は抑える
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
    st.subheader("🎉 完成＆ダウンロード")
    st.text_area("完成テキスト", st.session_state.final_text, height=300)

    if uploaded_template:
        try:
            wb = load_workbook(uploaded_template)
            ws = wb.active

            # A9セル以降をクリア
            for row in range(EXCEL_START_ROW, 100):
                ws[f"A{row}"].value = None

            # 分割して書き込み
            lines = split_text(st.session_state.final_text, CHARS_PER_LINE)
            
            for i, line in enumerate(lines):
                cell = ws[f"A{EXCEL_START_ROW + i}"]
                cell.value = line
                cell.alignment = Alignment(shrink_to_fit=True, wrap_text=False)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)

            st.download_button("📥 Excelをダウンロード", out, "感想文.xlsx", type="primary")
        except Exception as e:
            st.error(f"Excelエラー: {e}")
    else:
        st.warning("テンプレート(.xlsx)をアップロードしてください。")
