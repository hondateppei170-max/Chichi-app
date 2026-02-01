import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from PIL import Image
import io

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="致知読書感想文アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ (OCR強化版)")
st.caption("Step 1：画像解析 (Gemini) → Step 2：感想文執筆 (GPT-4o)")

# Excel書き込み設定
EXCEL_START_ROW = 9
CHARS_PER_LINE = 40

# ==========================================
# API設定
# ==========================================
try:
    # OpenAI (執筆用)
    openai_key = st.secrets.get("OPENAI_API_KEY")
    if not openai_key:
        st.warning("⚠️ OpenAI APIキーが設定されていません（secrets.tomlを確認してください）。")
    else:
        client = OpenAI(api_key=openai_key)

    # Google Gemini (画像解析用)
    google_key = st.secrets.get("GOOGLE_API_KEY")
    if not google_key:
        st.warning("⚠️ Google APIキーが設定されていません（secrets.tomlを確認してください）。")
    else:
        genai.configure(api_key=google_key)
    
except Exception as e:
    st.error(f"API設定エラー: {e}")
    st.stop()

# ==========================================
# 関数定義
# ==========================================
def split_text(text, chunk_size):
    """Excel用にテキストを指定文字数で分割"""
    if not text:
        return []
    clean_text = text.replace('\n', '　')
    return [clean_text[i:i+chunk_size] for i in range(0, len(clean_text), chunk_size)]

# ==========================================
# セッション状態
# ==========================================
if "extracted_text" not in st.session_state:
    st.session_state.extracted_text = ""
if "final_text" not in st.session_state:
    st.session_state.final_text = ""

# ==========================================
# サイドバー
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    uploaded_template = st.file_uploader("感想文フォーマット(.xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)

# ==========================================
# Step 1: 画像解析 (Gemini)
# ==========================================
st.header("Step 1. 記事画像の解析")
st.info("💡 モデルを使用して画像をOCR処理し、テキストデータ化します。")

# 3つの記事に対応するタブ
tab1, tab2, tab3 = st.tabs(["📂 メイン記事", "📂 記事2 (任意)", "📂 記事3 (任意)"])

files_dict = {}

with tab1:
    files_dict["main"] = st.file_uploader("メイン記事の画像", type=['png', 'jpg', 'jpeg', 'webp'], accept_multiple_files=True, key="u1")
with tab2:
    files_dict["sub1"] = st.file_uploader("記事2の画像", type=['png', 'jpg', 'jpeg', 'webp'], accept_multiple_files=True, key="u2")
with tab3:
    files_dict["sub2"] = st.file_uploader("記事3の画像", type=['png', 'jpg', 'jpeg', 'webp'], accept_multiple_files=True, key="u3")

total_files = sum([len(f) for f in files_dict.values() if f])

if total_files > 0:
    st.write(f"📁 合計 {total_files}枚の画像を読み込みました")

    if st.button("🔍 画像解析を開始 (OCR)", type="primary"):
        with st.spinner("Geminiが画像を読み込んでいます..."):
            try:
                gemini_inputs = []
                
                # ==========================================
                # プロンプト修正：OCR特化・位置情報付与
                # ==========================================
                system_prompt = """
                あなたは、雑誌『致知』の紙面を完璧に読み取る高精度OCRエンジンです。
                提供された全ての画像から、文字を一字一句漏らさず、ありのままに書き起こしてください。

                【目的】
                後続の処理でGPT-4oが記事を解析し、正確な「引用（掲載位置付き）」を作成するための元データを作成する。

                【厳守ルール】
                1. 完全な文字起こし（要約禁止）:
                   - 要約や省略は一切禁止。書いてある文字を一字一句正確に書き起こすこと。
                   - 縦書き（右上から左下）の文章の流れを正しく認識すること。
                   - インタビュー記事の場合、話し手の名前や「――」（ダッシュ）もそのまま記述すること。

                2. 位置情報のタグ付け（最重要）:
                   - 後で「1枚目 右段」と特定できるように、テキストの前に位置情報を付記すること。
                   - 画像ファイル名が判別できる場合は [ファイル名: xxx.jpg] と記載し、無理な場合は [画像N枚目] とする。
                   - 書式例：
                     [画像1枚目]
                     <タイトル> ...テキスト...
                     <リード文> ...テキスト...
                     <右段> ...テキスト...
                     <左段> ...テキスト...

                3. 記事ごとの区切り:
                   - 提供される画像は複数の記事に分かれている場合があるため、入力データの区切り指示（例：【ここからメイン記事】）に従ってセクションを分けること。

                4. 精度管理:
                   - 潰れて読めない文字は、勝手に創作せず「(判読不能)」と記述すること。
                   - 読めない部分があっても、読める部分は全て出力すること。
                """
                gemini_inputs.append(system_prompt)

                # 各タブの画像を処理
                article_labels = {
                    "main": "\n\n=== 【ここからメイン記事の画像】 ===\n", 
                    "sub1": "\n\n=== 【ここから記事2の画像】 ===\n", 
                    "sub2": "\n\n=== 【ここから記事3の画像】 ===\n"
                }

                for key, files in files_dict.items():
                    if files:
                        # ファイル名順ソート
                        files.sort(key=lambda x: x.name)
                        
                        gemini_inputs.append(article_labels[key])
                        
                        for img_file in files:
                            image = Image.open(img_file)
                            gemini_inputs.append(image)

                # Gemini モデル呼び出し
                # ※ gemini-3-flash がまだ利用できない場合は gemini-1.5-flash または gemini-2.0-flash-exp を指定してください
                model = genai.GenerativeModel('gemini-1.5-flash') 
                response = model.generate_content(gemini_inputs)

                st.session_state.extracted_text = response.text
                st.session_state.final_text = ""
                st.rerun()

            except Exception as e:
                st.error(f"Gemini解析エラー: {e}")

# ==========================================
# 解析結果の編集
# ==========================================
if st.session_state.extracted_text:
    st.markdown("---")
    st.subheader("📝 解析結果 (OCRデータ)")
    st.info("Step 2に進む前に、誤字や重要な箇所の欠落がないか確認・修正できます。")
    edited_text = st.text_area(
        "OCR結果編集エリア", 
        st.session_state.extracted_text, 
        height=500
    )
    st.session_state.extracted_text = edited_text

    # ==========================================
    # Step 2: 感想文作成 (OpenAI)
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の執筆 (GPT-4o)")

    if st.button("✍️ 税理士事務所員として感想文を書く"):
        if not st.session_state.extracted_text:
             st.error("解析データが空です。Step 1を実行してください。")
        else:
            with st.spinner("GPT-4oが執筆中..."):
                try:
                    # ==========================================
                    # 執筆プロンプト修正：OCRデータを活用する指示へ変更
                    # ==========================================
                    writer_prompt = f"""
                    あなたは税理士事務所の職員です。
                    以下の【OCR解析データ】は、雑誌『致知』の記事を文字起こししたものです。
                    この内容を元に、社内木鶏会用の読書感想文を作成してください。

                    【OCR解析データ】
                    {st.session_state.extracted_text}

                    【構成】
                    1. 記事の要約
                       - メイン記事の内容を中心に要約する。
                    
                    2. 印象に残った言葉（引用）
                       - 解析データ内の原文を引用する際は、必ず正確に記述すること。
                       - 引用部分の後に、（〇〇記事 〇枚目 右段より）のように、解析データにある位置情報を元に出典元を記載すること。
                       - 例：「～～～という言葉がありました（メイン記事 2枚目 左段より）」

                    3. 自分の業務への具体的な活かし方
                       - 税理士補助・監査業務などの視点で、記事の教えをどう実践するか具体的に書く。
                       - 精神論だけでなく、日々の顧客対応やミス防止、チームワークなどに結び付ける。

                    【執筆条件】
                    - 文字数：{target_length}文字前後
                    - 文体：「です・ます」調
                    - タイトルは不要。段落ごとに改行を入れること。
                    - 解析データに含まれていない内容（嘘の引用など）は絶対に創作しないこと。
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
    st.subheader("🎉 完成＆Excel出力")
    st.text_area("完成テキスト", st.session_state.final_text, height=300)

    if uploaded_template:
        try:
            wb = load_workbook(uploaded_template)
            ws = wb.active

            # A9セル以降クリア
            for row in range(EXCEL_START_ROW, 100):
                cell = ws[f"A{row}"]
                cell.value = None

            # 指定文字数で分割して書き込み
            lines = split_text(st.session_state.final_text, CHARS_PER_LINE)
            
            for i, line in enumerate(lines):
                cell = ws[f"A{EXCEL_START_ROW + i}"]
                cell.value = line
                # スタイル維持（折り返し設定など）
                current_alignment = cell.alignment
                cell.alignment = Alignment(
                    horizontal=current_alignment.horizontal,
                    vertical=current_alignment.vertical,
                    wrap_text=current_alignment.wrap_text,
                    shrink_to_fit=current_alignment.shrink_to_fit
                )

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)

            st.download_button("📥 Excelダウンロード", out, "致知感想文.xlsx", type="primary")
        except Exception as e:
            st.error(f"Excel処理エラー: {e}")
    else:
        st.warning("Excelテンプレートをアップロードすると、ダウンロードボタンが表示されます。")
