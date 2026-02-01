import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
from PIL import Image
import concurrent.futures

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(
    page_title="致知読書感想文アプリ v5.5",
    layout="wide",
    page_icon="📖"
)

# ==========================================
# 【ユーザー設定エリア】
# ==========================================
PAST_REVIEWS = """
（例：過去の感想文）
今月の致知を読んで、特に「逆境こそが人を育てる」という言葉が胸に刺さりました。
日々の税理士補助業務において、繁忙期にはつい愚痴が出そうになりますが、
それは自分の魂を磨く砥石なのだと気づかされました。
お客様の試算表を作る作業一つとっても、そこに魂を込めること。
それがプロフェッショナルとしての流儀だと感じます。
"""

EXCEL_START_ROW = 9
CHARS_PER_LINE = 40

# ==========================================
# API設定 (サイドバー)
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    
    openai_key = st.secrets.get("OPENAI_API_KEY")
    if not openai_key:
        openai_key = st.text_input("OpenAI API Key", type="password")
    
    google_key = st.secrets.get("GOOGLE_API_KEY")
    if not google_key:
        google_key = st.text_input("Google API Key", type="password")

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
        st.session_state.clear()
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
# 【重要】テキストエリアを強制リフレッシュするためのカウンタ
if "rewrite_count" not in st.session_state:
    st.session_state.rewrite_count = 0

# ==========================================
# 関数定義
# ==========================================
def split_text(text, chunk_size):
    if not text: return []
    clean_text = text.replace('\n', '　')
    return [clean_text[i:i+chunk_size] for i in range(0, len(clean_text), chunk_size)]

def process_ocr_task_safe(label, pil_images, model_id):
    if not pil_images: return ""
    try:
        gemini_inputs = []
        system_prompt = (
            "あなたは高精度なOCRエンジンです。雑誌『致知』の紙面を読み取ります。\n"
            "レイアウト（段組み）を認識し、記事のブロックごとに、右から左へ縦書きの流れを汲んで文字起こしをしてください。\n"
            "出力形式: [画像番号] <本文>..."
        )
        gemini_inputs.append(system_prompt)
        for i, img in enumerate(pil_images):
            gemini_inputs.append(f"\n\n[画像{i+1}枚目]\n")
            gemini_inputs.append(img)
        model = genai.GenerativeModel(model_id)
        response = model.generate_content(gemini_inputs)
        return response.text
    except Exception as e:
        return f"[エラー: {label}の解析失敗: {e}]"

def generate_draft(article_text, chat_context, target_len):
    if not client: return "エラー: OpenAI APIキーが必要です。"

    # プロンプトの切り替え
    if not chat_context:
        # 初稿モード
        system_prompt = (
            "あなたは税理士事務所の職員です。雑誌『致知』の読書感想文の【初稿】を作成します。\n"
            "過去の文体サンプルを模倣し、記事を要約してください。\n"
            "**重要：まだ具体的な体験談は入力されていません。「日々の業務において〜」等の一般的な表現で留めてください。創作は厳禁です。**"
        )
        user_content = (
            f"【記事データ】\n{article_text}\n\n"
            f"【文体サンプル】\n{PAST_REVIEWS}\n\n"
            f"【文字数】{target_len}文字前後"
        )
    else:
        # 書き直しモード（強力な反映指示）
        system_prompt = (
            "あなたはプロのライターです。読書感想文の【エピソード差し替え】を行います。\n"
            "現在あるドラフトの「一般的な業務の話」部分を削除し、\n"
            "**以下のチャットログにある『具体的な体験談』に完全に書き換えてください。**\n"
            "チャットで語られた内容（いつ、誰が、どうした）が含まれていなければ失敗とみなします。"
        )
        user_content = (
            f"【最優先：組み込むべきユーザーのエピソード】\n"
            f"--------------------------------------------------\n"
            f"{chat_context}\n"
            f"--------------------------------------------------\n"
            f"↑この内容を感想文のメインパート（全体の6割）として展開してください。\n\n"
            f"【元記事】\n{article_text}\n\n"
            f"【文字数】{target_len}文字前後"
        )
    
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": user_content}],
        temperature=0.7
    )
    return response.choices[0].message.content

# ==========================================
# メイン画面構成
# ==========================================
st.title("📖 致知読書感想文アプリ v5.5 (強制更新版)")
st.caption("Step 1: OCR → Step 2: 記事選択・執筆 → Step 3: Excel出力")

tab1, tab2, tab3 = st.tabs(["1️⃣ 画像解析", "2️⃣ 記事選択 & 執筆", "3️⃣ Excel出力"])

# ------------------------------------------------------------------
# Tab 1: OCR
# ------------------------------------------------------------------
with tab1:
    st.subheader("Step 1. 記事画像の読み込み")
    col1, col2, col3 = st.columns(3)
    with col1:
        files_main = st.file_uploader("メイン記事", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="f1")
    with col2:
        files_sub1 = st.file_uploader("記事2", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="f2")
    with col3:
        files_sub2 = st.file_uploader("記事3", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="f3")

    if st.button("🚀 解析スタート", type="primary"):
        if not (files_main or files_sub1 or files_sub2):
            st.error("画像を選択してください。")
        elif not google_key:
            st.error("Google APIキーが必要です。")
        else:
            with st.spinner("解析中..."):
                try:
                    images_main = [Image.open(f).convert("RGB") for f in files_main] if files_main else []
                    images_sub1 = [Image.open(f).convert("RGB") for f in files_sub1] if files_sub1 else []
                    images_sub2 = [Image.open(f).convert("RGB") for f in files_sub2] if files_sub2 else []

                    with concurrent.futures.ThreadPoolExecutor() as executor:
                        f1 = executor.submit(process_ocr_task_safe, "メイン", images_main, model_main)
                        f2 = executor.submit(process_ocr_task_safe, "記事2", images_sub1, model_sub)
                        f3 = executor.submit(process_ocr_task_safe, "記事3", images_sub2, model_sub)
                        st.session_state.ocr_results["main"] = f1.result()
                        st.session_state.ocr_results["sub1"] = f2.result()
                        st.session_state.ocr_results["sub2"] = f3.result()
                    st.success("解析完了！")
                except Exception as e:
                    st.error(f"エラー: {e}")

    with st.expander("OCR結果詳細"):
        st.text_area("Main", st.session_state.ocr_results["main"], height=100)

# ------------------------------------------------------------------
# Tab 2: 執筆 & 壁打ち (ここが修正の核心)
# ------------------------------------------------------------------
with tab2:
    st.subheader("Step 2. 執筆 & 壁打ち")
    
    # 記事選択
    options = {k: v for k, v in st.session_state.ocr_results.items() if len(v) > 10}
    map_label = {"main": "メイン記事", "sub1": "記事2", "sub2": "記事3"}
    
    if not options:
        st.warning("まずはTab 1でOCRを実行してください。")
        selected_text = ""
    else:
        sel = st.radio("執筆対象", list(options.keys()), format_func=lambda x: map_label[x], horizontal=True)
        selected_text = options[sel]

    st.markdown("---")
    
    # 左右カラム定義
    col_draft, col_chat = st.columns([1, 1])

    # ------------------------------------------------
    # 左カラム：感想文ドラフト
    # ------------------------------------------------
    with col_draft:
        st.markdown("### 📝 感想文")
        
        # 初稿作成
        if st.button("🚀 初稿を作成 (まだエピソードなし)"):
            if not client:
                st.error("OpenAI APIキーが必要です。")
            else:
                with st.spinner("初稿作成中..."):
                    st.session_state.chat_history = [] # 履歴リセット
                    draft = generate_draft(selected_text, None, target_length)
                    st.session_state.current_draft = draft
                    st.session_state.rewrite_count += 1 # 強制更新用カウントアップ
                    
                    # 最初のAIメッセージ
                    st.session_state.chat_history.append({
                        "role": "assistant",
                        "content": "初稿を作りました。\n**この記事に関連して、あなたの業務での具体的な体験談（成功・失敗）を右のチャットで教えてください。**"
                    })
                    st.rerun()

        # 書き直しボタン
        if st.button("🔄 チャット内容を反映して書き直す", type="primary"):
            if len(st.session_state.chat_history) <= 1:
                st.warning("右側のチャットでエピソードを入力してください。")
            else:
                with st.spinner("チャットのエピソードを反映中..."):
                    # チャット履歴を全部渡す
                    chat_log = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.chat_history])
                    new_draft = generate_draft(selected_text, chat_log, target_length)
                    
                    st.session_state.current_draft = new_draft
                    st.session_state.rewrite_count += 1 # 【重要】これでテキストエリアが生まれ変わる
                    st.success("反映完了！")
                    st.rerun()

        # ドラフト表示エリア
        if st.session_state.current_draft:
            # keyを動的に変えることで、Streamlitに「新しいウィジェットだ」と認識させ、強制的にvalueを読み込ませる
            dynamic_key = f"draft_area_{st.session_state.rewrite_count}"
            
            st.text_area(
                "現在の原稿", 
                value=st.session_state.current_draft, 
                height=600, 
                key=dynamic_key
            )

    # ------------------------------------------------
    # 右カラム：壁打ちチャット (常に表示)
    # ------------------------------------------------
    with col_chat:
        st.markdown("### 💬 エピソード深掘りチャット")
        chat_container = st.container(height=500)
        
        # 履歴表示
        for msg in st.session_state.chat_history:
            with chat_container.chat_message(msg["role"]):
                st.markdown(msg["content"])

        # 入力フォーム
        if prompt := st.chat_input("体験談を入力..."):
            if not selected_text:
                st.error("先に初稿を作成してください。")
            elif not client:
                st.error("OpenAI APIキーが必要です。")
            else:
                # ユーザーの入力を追加
                st.session_state.chat_history.append({"role": "user", "content": prompt})
                with chat_container.chat_message("user"):
                    st.markdown(prompt)

                # AIの返答生成
                with chat_container.chat_message("assistant"):
                    with st.spinner("考え中..."):
                        sys_msg = f"あなたは編集者です。以下の記事: {selected_text[:300]}... を踏まえ、ユーザーからより深いエピソード（いつ、誰が、どうした）を引き出す質問をしてください。"
                        msgs = [{"role": "system", "content": sys_msg}] + st.session_state.chat_history
                        res = client.chat.completions.create(model="gpt-4o", messages=msgs)
                        ai_msg = res.choices[0].message.content
                
                st.markdown(ai_msg)
                st.session_state.chat_history.append({"role": "assistant", "content": ai_msg})

# ------------------------------------------------------------------
# Tab 3: Excel出力
# ------------------------------------------------------------------
with tab3:
    st.subheader("Step 3. Excel出力")
    if st.session_state.current_draft and uploaded_template:
        if st.button("📥 Excelダウンロード"):
            try:
                wb = load_workbook(uploaded_template)
                ws = wb.active
                for r in range(EXCEL_START_ROW, 100): ws[f"A{r}"].value = None
                lines = split_text(st.session_state.current_draft, CHARS_PER_LINE)
                for i, line in enumerate(lines):
                    cell = ws[f"A{EXCEL_START_ROW+i}"]
                    cell.value = line
                    cell.alignment = Alignment(wrap_text=False, shrink_to_fit=False, horizontal='left')
                out = io.BytesIO()
                wb.save(out)
                out.seek(0)
                st.download_button("Excel保存", out, "感想文.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.success("完了！")
            except Exception as e:
                st.error(f"エラー: {e}")
