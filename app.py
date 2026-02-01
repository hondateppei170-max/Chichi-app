import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
from PIL import Image
import concurrent.futures

# ==========================================
# ページ設定 (必ず一番最初に書く)
# ==========================================
st.set_page_config(page_title="致知読書感想文アプリ v4.2 (順序強制版)", layout="wide", page_icon="📖")

# ==========================================
# 【ユーザー設定エリア: 過去の文体学習】
# ==========================================
PAST_REVIEWS = """
（例：過去の感想文）
今月の致知を読んで、特に「逆境こそが人を育てる」という言葉が胸に刺さりました。
日々の税理士補助業務において、繁忙期にはつい愚痴が出そうになりますが、
それは自分の魂を磨く砥石なのだと気づかされました。
お客様の試算表を作る作業一つとっても、そこに魂を込めること。
それがプロフェッショナルとしての流儀だと感じます。
"""

# Excel書き込み設定
EXCEL_START_ROW = 9
CHARS_PER_LINE = 40

# ==========================================
# API設定
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    
    # API Key入力（secretsになければ入力欄表示）
    openai_key = st.secrets.get("OPENAI_API_KEY")
    if not openai_key:
        openai_key = st.text_input("OpenAI API Key", type="password")
    
    google_key = st.secrets.get("GOOGLE_API_KEY")
    if not google_key:
        google_key = st.text_input("Google API Key", type="password")

    # Client初期化
    client = None
    if openai_key:
        try:
            client = OpenAI(api_key=openai_key)
        except:
            st.error("OpenAIキーが無効です")

    if google_key:
        try:
            genai.configure(api_key=google_key)
        except:
            st.error("Googleキーが無効です")

    st.markdown("---")
    uploaded_template = st.file_uploader("感想文フォーマット(.xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)
    
    st.markdown("---")
    st.caption("🔧 OCRモデル設定")
    model_main = st.text_input("メインModel ID", value="gemini-3-flash-preview")
    model_sub = st.text_input("サブModel ID", value="gemini-2.0-flash-lite-preview-02-05")

    if st.button("🗑️ リセット"):
        for key in st.session_state.keys():
            del st.session_state[key]
        st.rerun()

# ==========================================
# セッション状態の初期化
# ==========================================
if "ocr_results" not in st.session_state:
    st.session_state.ocr_results = {"main": "", "sub1": "", "sub2": ""}
if "current_draft" not in st.session_state:
    st.session_state.current_draft = ""
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "selected_article_key" not in st.session_state:
    st.session_state.selected_article_key = "main"

# ==========================================
# 関数定義
# ==========================================
def split_text(text, chunk_size):
    if not text: return []
    clean_text = text.replace('\n', '　')
    return [clean_text[i:i+chunk_size] for i in range(0, len(clean_text), chunk_size)]

def process_ocr_task_safe(label, pil_images, model_id):
    """
    【修正版】並列処理用OCR関数
    画像を物理的に「上半分」と「下半分」に切り分けてからAIに渡すことで、
    強制的に「上段→下段」の順序で読ませる。
    """
    if not pil_images:
        return ""
    
    try:
        gemini_inputs = []
        # プロンプト修正：分割された画像が順番に来ることを伝える
        system_prompt = (
            "あなたはOCRエンジンです。\n"
            "これから雑誌『致知』のページを「上半分」と「下半分」に分割した画像が順番に送られます。\n"
            "送られてきた画像の順番通りに（まず上段部分、次に下段部分）、文字を書き起こしてください。\n"
            "縦書きの文章は、右行から左行へ読んでください。"
        )
        gemini_inputs.append(system_prompt)
        
        # 【重要】画像を物理的に上下分割してリストに追加
        for i, img in enumerate(pil_images):
            width, height = img.size
            
            # 上半分 (Top Half)
            top_half = img.crop((0, 0, width, height // 2))
            # 下半分 (Bottom Half)
            bottom_half = img.crop((0, height // 2, width, height))
            
            # 順番通りに追加 (これでAIは上から読むしかなくなる)
            gemini_inputs.append(f"\n\n[画像{i+1}枚目：上段エリア]\n")
            gemini_inputs.append(top_half)
            gemini_inputs.append(f"\n\n[画像{i+1}枚目：下段エリア]\n")
            gemini_inputs.append(bottom_half)
        
        # モデル実行
        model = genai.GenerativeModel(model_id)
        response = model.generate_content(gemini_inputs)
        return response.text
        
    except Exception as e:
        return f"[エラー: {label}の解析失敗: {e}]"

def generate_draft(article_text, chat_context, target_len):
    if not client:
        return "エラー: OpenAI APIキーが設定されていません。"

    system_prompt = (
        "あなたは税理士事務所の職員です。\n"
        "これから雑誌『致知』の読書感想文（社内木鶏会用）を作成します。\n"
        "以下の【ユーザーの過去の感想文】を分析し、"
        "**「文体」「書き出しの癖」「精神的な熱量」「業務（巡回監査・決算など）への結びつけ方」**を模倣してください。"
    )
    user_content = (
        f"【今回選択した記事のOCRデータ】\n{article_text}\n\n"
        f"【ユーザーの過去の感想文（スタイル見本）】\n{PAST_REVIEWS}\n\n"
        f"【打ち合わせ内容】\n{chat_context}\n\n"
        "【執筆条件】\n"
        f"- 文字数：{target_len}文字前後\n"
        "- 文体：「です・ます」調\n"
        "- 段落ごとに改行を入れること。\n"
        "- 構成：①記事の引用 ②自分の業務エピソード ③今後の決意"
    )
    
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": user_content}],
        temperature=0.7
    )
    return response.choices[0].message.content

# ==========================================
# メイン画面
# ==========================================
st.title("📖 致知読書感想文アプリ v4.2 (順序強制版)")
st.caption("Step
