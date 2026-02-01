import streamlit as st
import google.generativeai as genai
from openai import OpenAI
import openpyxl
from openpyxl.styles import Alignment
import io
from PIL import Image

# --- 設定 ---
st.set_page_config(page_title="致知読書感想作成アプリ", layout="wide")
st.title("致知 読書感想文作成支援 (Gemini × GPT-4o)")

# サイドバーでAPIキー設定
with st.sidebar:
    gemini_key = st.text_input("Google API Key", type="password")
    openai_key = st.text_input("OpenAI API Key", type="password")
    st.info("Geminiで高精度OCRを行い、GPT-4oで感想文を生成します。")

if not gemini_key or not openai_key:
    st.warning("両方のAPIキーを入力してください。")
    st.stop()

# 各クライアント初期化
genai.configure(api_key=gemini_key)
gemini_model = genai.GenerativeModel('gemini-1.5-flash')
client = OpenAI(api_key=openai_key)

# --- Step 1: 画像読み込みとOCR ---
st.header("Step 1: 記事の解析と事実抽出")
uploaded_files = st.file_uploader("『致知』の画像をアップロード（15枚以上対応）", 
                                  type=['png', 'jpg', 'jpeg'], 
                                  accept_multiple_files=True)

if uploaded_files:
    # ファイル名順にソート
    sorted_files = sorted(uploaded_files, key=lambda x: x.name)
    
    if st.button("記事を解析する"):
        combined_ocr_text = ""
        progress_bar = st.progress(0)
        
        for i, file in enumerate(sorted_files):
            img = Image.open(file)
            # Geminiによる高精度OCR & 構造化抽出
            prompt = """
            この画像は雑誌『致知』のページです。
            1. 記事の内容を正確にテキスト化してください。
            2. 重要な文章を抜き出し、必ずその場所（例：〇枚目 右段 〇行目付近）を付記してください。
            3. 創作は一切せず、書かれている事実のみを抽出してください。
            """
            response = gemini_model.generate_content([prompt, img])
            combined_ocr_text += f"\n\n--- {file.name} の解析結果 ---\n{response.text}"
            progress_bar.progress((i + 1) / len(sorted_files))
        
        st.session_state['ocr_result'] = combined_ocr_text

    if 'ocr_result' in st.session_state:
        edited_text = st.text_area("抽出結果の確認・編集（この内容を元に感想文を作ります）", 
                                    st.session_state['ocr_result'], height=400)
        st.session_state['final_context'] = edited_text

# --- Step 2: 感想文の執筆 ---
if 'final_context' in st.session_state:
    st.header("Step 2: 感想文の執筆")
    template_file = st.file_uploader("Excelテンプレート（.xlsx）をアップロード", type=['xlsx'])
    
    if st.button("感想文を生成する") and template_file:
        # GPT-4oによる執筆
        system_prompt = """
        あなたは税理士事務所の職員です。雑誌『致知』を読み、社内木鶏会用の感想文を作成します。
        与えられたテキストのみを根拠とし、以下の構成で作成してください。
        ①記事の要約
        ②印象に残った言葉（正確な引用と場所を記載）
        ③自分の業務（税理士事務所職員としての視点）への活かし方
        
        【ルール】
        - 創作や記事にない情報の追加は厳禁。
        - タイトルは不要。
        - 段落ごとの改行を入れること。
        """
        
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": st.session_state['final_context']}
            ],
            temperature=0.0
        )
        
        draft_text = response.choices[0].message.content
        st.subheader("生成された感想文")
        st.write(draft_text)
        
        # --- Excel出力処理 ---
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active
        
        # 1行40文字で分割
        flat_text = draft_text.replace('\n', ' ')
        lines = [flat_text[i:i+40] for i in range(0, len(flat_text), 40)]
        
        # A9セルから順に書き込み
        for idx, line in enumerate(lines):
            cell = ws.cell(row=9 + idx, column=1)
            cell.value = line
            cell.alignment = Alignment(shrink_to_fit=True, vertical='center')
        
        # 保存準備
        output = io.BytesIO()
        wb.save(output)
        st.download_button(label="Excelファイルをダウンロード", 
                           data=output.getvalue(), 
                           file_name="致知感想文_出力.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
