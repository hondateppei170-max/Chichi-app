import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
from PIL import Image

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="致知読書感想文アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ (Gemini 3.0版)")
st.caption("Step 1：画像解析 (Gemini 3 Flash) → Step 2：感想文執筆 (GPT-4o)")

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
        st.warning("⚠️ OpenAI APIキーが設定されていません。")
    else:
        client = OpenAI(api_key=openai_key)

    # Google Gemini (画像解析用)
    google_key = st.secrets.get("GOOGLE_API_KEY")
    if not google_key:
        st.warning("⚠️ Google APIキーが設定されていません。")
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
# サイドバー設定
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    uploaded_template = st.file_uploader("感想文フォーマット(.xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)
    
    st.markdown("---")
    st.caption("🔧 モデル設定")
    # デフォルト値を gemini-3-flash に設定
    model_id_input = st.text_input("GeminiモデルID", value="gemini-3-flash")
    st.caption("※Google AI Studio等で確認できるモデル名を入力")

# ==========================================
# Step 1: 画像解析 (Gemini 3 Flash)
# ==========================================
st.header("Step 1. 記事画像の解析")
st.info(f"💡 指定モデル「{model_id_input}」を使用して画像をOCR処理します。")

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
        with st.spinner(f"Gemini ({model_id_input}) が画像を読み込んでいます..."):
            try:
                gemini_inputs = []
                
                # システムプロンプト（OCR特化）
                # 複雑なインデントエラーを避けるため、変数で定義
                system_prompt_text = (
                    "あなたは、雑誌『致知』の紙面を完璧に読み取る高精度OCRエンジンです。\n"
                    "提供された全ての画像から、文字を一字一句漏らさず、ありのままに書き起こしてください。\n\n"
                    "【目的】\n"
                    "後続の処理でGPT-4oが記事を解析し、正確な「引用（掲載位置付き）」を作成するための元データを作成する。\n\n"
                    "【厳守ルール】\n"
                    "1. 完全な文字起こし（要約禁止）:\n"
                    "   - 要約や省略は一切禁止。書いてある文字を一字一句正確に書き起こすこと。\n"
                    "   - 縦書き（右上から左下）の文章の流れを正しく認識すること。\n\n"
                    "2. 位置情報のタグ付け（最重要）:\n"
                    "   - 後で「1枚目 右段」と特定できるように、テキストの前に位置情報を付記すること。\n"
                    "   - 画像ファイル名が判別できる場合は [ファイル名: xxx.jpg] と記載し、無理な場合は [画像N枚目] とする。\n\n"
                    "3. 記事ごとの区切り:\n"
                    "   - 提供される画像は複数の記事に分かれている場合があるため、入力データの区切り指示に従ってセクションを分けること。"
                )
                gemini_inputs.append(system_prompt_text)

                # 画像データの準備
                for key, files in files_dict.items():
                    if files:
                        # 名前順にソート
                        files.sort(key=lambda x: x.name)
                        
                        # セクション区切りテキストを追加
                        if key == "main":
                            gemini_inputs.append("\n\n=== 【ここからメイン記事の画像】 ===\n")
                        elif key == "sub1":
                            gemini_inputs.append("\n\n=== 【ここから記事2の画像】 ===\n")
                        elif key == "sub2":
                            gemini_inputs.append("\n\n=== 【ここから記事3の画像】 ===\n")
                        
                        # 画像を追加（エラー対策済み）
                        for img_file in files:
                            try:
                                # 【重要】ポインタを先頭に戻す
                                img_file.seek(0)
                                
                                image = Image.open(img_file)
                                
                                # RGB変換（AlphaチャンネルやCMYK対策）
                                if image.mode != "RGB":
                                    image = image.convert("RGB")
                                    
                                gemini_inputs.append(image)
                                
                            except Exception as img_err:
                                st.error(f"画像読み込みエラー: {img_file.name} \n詳細: {img_err}")
                                # 読み込めない画像があっても停止せず、次の画像へ進む
                                continue

                # Gemini モデル呼び出し
                try:
                    model = genai.GenerativeModel(model_id_input)
                    response = model.generate_content(gemini_inputs)
                    
                    st.session_state.extracted_text = response.text
                    st.session_state.final_text = ""
                    st.success("✅ 解析完了")
                    st.rerun()

                except Exception as e_model:
                    st.error(f"モデル「{model_id_input}」での解析に失敗しました。")
                    st.error(f"詳細エラー: {e_model}")
                    st.markdown("---")
                    st.warning("利用可能なモデル一覧（参考）:")
                    try:
                        available_models = []
                        for m in genai.list_models():
                            if 'generateContent' in m.supported_generation_methods:
                                available_models.append(m.name)
                        st.code("\n".join(available_models))
                    except:
                        st.write("(モデル一覧の取得に失敗しました)")
                    st.stop()

            except Exception as e:
                st.error(f"システムエラー: {e}")

# ==========================================
# 解析結果の編集
# ==========================================
if st.session_state.extracted_text:
    st.markdown("---")
    st.subheader("📝 解析結果 (OCRデータ)")
    
    edited_text = st.text_area(
        label="OCR結果編集エリア",
        value=st.session_state.extracted_text,
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
                    # 執筆プロンプト（安全な結合）
                    ocr_data = st.session_state.extracted_text
                    
                    writer_prompt = (
                        "あなたは税理士事務所の職員です。\n"
                        "以下の【OCR解析データ】は、雑誌『致知』の記事を文字起こししたものです。\n"
                        "この内容を元に、社内木鶏会用の読書感想文を作成してください。\n\n"
                        "【OCR解析データ】\n"
                        f"{ocr_data}\n\n"
                        "【構成】\n"
                        "1. 記事の要約\n"
                        "   - メイン記事の内容を中心に要約する。\n"
                        "2. 印象に残った言葉（引用）\n"
                        "   - 解析データ内の原文を引用する際は、必ず正確に記述すること。\n"
                        "   - 引用部分の後に、（〇〇記事 〇枚目 右段より）のように、解析データにある位置情報を元に出典元を記載すること。\n"
                        "3. 自分の業務への具体的な活かし方\n"
                        "   - 税理士補助・監査業務などの視点で、記事の教えをどう実践するか具体的に書く。\n\n"
                        "【執筆条件】\n"
                        f"- 文字数：{target_length}文字前後\n"
                        "- 文体：「です・ます」調\n"
                        "- タイトル不要。段落ごとに改行。\n"
                        "- 解析データにない内容は創作しないこと。"
                    )

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
            
            # 書き込み前のクリア
            for row in range(EXCEL_START_ROW, 100):
                ws[f"A{row}"].value = None
            
            # 分割書き込み
            lines = split_text(st.session_state.final_text, CHARS_PER_LINE)
            for i, line in enumerate(lines):
                cell = ws[f"A{EXCEL_START_ROW + i}"]
                cell.value = line
                # スタイル調整
                cell.alignment = Alignment(wrap_text=False, shrink_to_fit=True)

            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            
            st.download_button("📥 Excelダウンロード", out, "致知感想文.xlsx", type="primary")
            
        except Exception as e:
            st.error(f"Excel処理エラー: {e}")
