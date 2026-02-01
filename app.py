import streamlit as st
import google.generativeai as genai
from openai import OpenAI
import openpyxl
from openpyxl.styles import Alignment
import io
from PIL import Image

# --- アプリ設定 ---
st.set_page_config(page_title="致知・木鶏会感想文作成", layout="wide")

# APIキーの取得（サイドバー）
with st.sidebar:
    st.header("API設定")
    gemini_key = st.text_input("Google Gemini API Key", type="password")
    openai_key = st.text_input("OpenAI API Key (GPT-4o)", type="password")
    
    st.divider()
    st.write("【手順】")
    st.write("1. 画像を全枚数アップロード")
    st.write("2. GeminiでOCR解析を実行")
    st.write("3. 抽出結果を確認・修正")
    st.write("4. GPT-4oで感想文を生成・Excel出力")

# クライアントの初期化関数
def init_apis(g_key, o_key):
    genai.configure(api_key=g_key)
    gemini_model = genai.GenerativeModel('gemini-1.5-flash')
    openai_client = OpenAI(api_key=o_key)
    return gemini_model, openai_client

# --- メイン UI ---
st.title("致知 読書感想文作成アプリ")

if not gemini_key or not openai_key:
    st.warning("サイドバーから両方のAPIキーを入力してください。")
    st.stop()

gemini_model, client = init_apis(gemini_key, openai_key)

# --- Step 1: 画像読み込みとGeminiによるOCR ---
st.header("Step 1: Geminiによる高精度OCR解析")
uploaded_files = st.file_uploader(
    "『致知』の誌面画像をアップロードしてください（15枚以上一括可）", 
    type=['png', 'jpg', 'jpeg'], 
    accept_multiple_files=True
)

if uploaded_files:
    # ファイル名でソート (IMG_001, IMG_002...)
    sorted_files = sorted(uploaded_files, key=lambda x: x.name)
    
    if st.button("全画像を解析してテキストを抽出する"):
        combined_text = ""
        progress_text = "画像を解析中..."
        my_bar = st.progress(0, text=progress_text)
        
        for i, file in enumerate(sorted_files):
            img = Image.open(file)
            
            # GeminiへのOCR指示（厳格な事実抽出）
            prompt = f"""
            これは雑誌『致知』の{i+1}枚目の画像です。
            【厳守事項】
            1. 誌面の文字をすべて正確に書き起こしてください。
            2. 重要な文章を抜き出し、「{i+1}枚目 右段 〇行目付近」のように場所を特定して記載してください。
            3. あなたの感想や解釈は一切含めず、紙面に書かれている内容のみを出力してください。
            """
            
            try:
                response = gemini_model.generate_content([prompt, img])
                combined_text += f"\n\n=== {file.name} (解析結果) ===\n{response.text}\n"
            except Exception as e:
                st.error(f"{file.name}の解析中にエラーが発生しました: {e}")
            
            my_bar.progress((i + 1) / len(sorted_files), text=f"{file.name} を解析完了")
        
        st.session_state['ocr_raw_data'] = combined_text

    if 'ocr_raw_data' in st.session_state:
        st.subheader("抽出結果の確認・編集")
        st.info("GPT-4oに渡す前に、誤字脱字や不要な情報をここで修正できます。")
        final_input = st.text_area("OCR抽出テキスト", st.session_state['ocr_raw_data'], height=400)
        st.session_state['processed_text'] = final_input

# --- Step 2: GPT-4oによる執筆とExcel出力 ---
if 'processed_text' in st.session_state:
    st.divider()
    st.header("Step 2: GPT-4oによる感想文作成")
    
    template_file = st.file_uploader("Excelテンプレート（.xlsx）をアップロード", type=['xlsx'])
    
    if st.button("感想文を執筆してExcelを作成") and template_file:
        with st.spinner("GPT-4oが執筆中..."):
            system_prompt = """
            あなたは税理士事務所の職員です。
            提示された『致知』の記事テキストを元に、社内木鶏会用の感想文を作成してください。
            
            【執筆ルール】
            - 記事にない事実は絶対に書かない。
            - 構成：
              ①記事の要約（核心を突いた内容）
              ②印象に残った言葉（「～」といった形で正確に引用し、出典の場所も明記）
              ③自分の業務（税理士事務所）への具体的な活かし方
            - トーン：誠実、謙虚、プロフェッショナル。
            - 改行は段落の切り替わり人のみ。タイトルなどは不要。
            """
            
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": st.session_state['processed_text']}
                ],
                temperature=0.0
            )
            
            essay_result = response.choices[0].message.content
            st.subheader("完成した感想文")
            st.write(essay_result)
            
            # Excel書き込み
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active
            
            # 1行40文字で分割するロジック
            clean_text = essay_result.replace('\n', ' ')
            max_chars = 40
            lines = [clean_text[i:i+max_chars] for i in range(0, len(clean_text), max_chars)]
            
            # A9セルから順に書き込み
            for r_idx, text_line in enumerate(lines):
                target_row = 9 + r_idx
                cell = ws.cell(row=target_row, column=1)
                cell.value = text_line
                cell.alignment = Alignment(shrink_to_fit=True, vertical='center')
            
            # ダウンロード
            output = io.BytesIO()
            wb.save(output)
            st.download_button(
                label="Excelファイルをダウンロード",
                data=output.getvalue(),
                file_name="社内木鶏会_感想文.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
