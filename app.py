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
st.set_page_config(page_title="致知読書感想文アプリ v4.1", layout="wide", page_icon="📖")

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
# API設定 (エラー回避のための安全策)
# ==========================================
# サイドバーでAPIキーを確認・入力できるようにする
with st.sidebar:
    st.header("⚙️ 設定")
    
    # OpenAI Key
    openai_key = st.secrets.get("OPENAI_API_KEY")
    if not openai_key:
        openai_key = st.text_input("OpenAI API Key", type="password")
    
    # Google Key
    google_key = st.secrets.get("GOOGLE_API_KEY")
    if not google_key:
        google_key = st.text_input("Google API Key", type="password")

    # 設定反映
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
    ファイルオブジェクトではなく、既にロードされたPIL画像を受け取ることでエラーを防ぐ
    """
    if not pil_images:
        return ""
    
    try:
        gemini_inputs = []
        system_prompt = (
            "あなたはOCRエンジンです。雑誌『致知』の「上下2段組み」画像を読み取ります。\n"
            "【厳守ルール】\n"
            "1. 画像を「上段」と「下段」に分けて認識する。\n"
            "2. まず【上段】の文章を右から左へ読む。\n"
            "3. 次に【下段】の文章を右から左へ読む。\n"
            "4. ※絶対に左側の段を上から下へ一気に読まないこと。\n"
            "5. 出力形式: [画像番号] <上段>... <下段>..."
        )
        gemini_inputs.append(system_prompt)
        
        # 既に画像データになっているのでそのまま追加
        gemini_inputs.extend(pil_images)
        
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
st.title("📖 致知読書感想文作成アプリ v4.1 (修正版)")
st.caption("Step 1: 並列OCR → Step 2: 記事選択・執筆 → Step 3: Excel出力")

tab1, tab2, tab3 = st.tabs(["1️⃣ 画像解析 (並列処理)", "2️⃣ 記事選択 & 執筆", "3️⃣ Excel出力"])

# ------------------------------------------------------------------
# Tab 1: 並列OCR処理
# ------------------------------------------------------------------
with tab1:
    st.subheader("Step 1. 記事画像の読み込み")
    st.info("※エラー防止のため、画像はメモリ上で処理してから並列解析にかけます。")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("#### 📂 メイン記事")
        files_main = st.file_uploader("画像を選択", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="f1")
    with col2:
        st.markdown("#### 📂 記事2")
        files_sub1 = st.file_uploader("画像を選択", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="f2")
    with col3:
        st.markdown("#### 📂 記事3")
        files_sub2 = st.file_uploader("画像を選択", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="f3")

    if st.button("🚀 全記事を一括解析 (並列スタート)", type="primary"):
        if not (files_main or files_sub1 or files_sub2):
            st.error("画像が選択されていません。")
        elif not google_key:
            st.error("Google APIキーが設定されていません。")
        else:
            with st.spinner("画像を読み込んで解析中..."):
                # 【修正点】スレッドに渡す前にメインスレッドで画像をPIL形式に変換する
                # これにより "ValueError: I/O operation on closed file" を防ぐ
                try:
                    images_main = [Image.open(f).convert("RGB") for f in files_main] if files_main else []
                    images_sub1 = [Image.open(f).convert("RGB") for f in files_sub1] if files_sub1 else []
                    images_sub2 = [Image.open(f).convert("RGB") for f in files_sub2] if files_sub2 else []

                    # 並列処理の実行
                    with concurrent.futures.ThreadPoolExecutor() as executor:
                        future_main = executor.submit(process_ocr_task_safe, "メイン記事", images_main, model_main)
                        future_sub1 = executor.submit(process_ocr_task_safe, "記事2", images_sub1, model_sub)
                        future_sub2 = executor.submit(process_ocr_task_safe, "記事3", images_sub2, model_sub)
                        
                        st.session_state.ocr_results["main"] = future_main.result()
                        st.session_state.ocr_results["sub1"] = future_sub1.result()
                        st.session_state.ocr_results["sub2"] = future_sub2.result()
                    
                    st.success("✅ 解析完了！ '2️⃣ 記事選択 & 執筆' タブへ進んでください。")
                except Exception as e:
                    st.error(f"予期せぬエラーが発生しました: {e}")

    with st.expander("OCR解析結果を確認する"):
        st.text_area("Main", st.session_state.ocr_results["main"], height=100)
        st.text_area("Sub1", st.session_state.ocr_results["sub1"], height=100)
        st.text_area("Sub2", st.session_state.ocr_results["sub2"], height=100)

# ------------------------------------------------------------------
# Tab 2: 記事選択 & 執筆 & 壁打ち
# ------------------------------------------------------------------
with tab2:
    st.subheader("Step 2. 執筆対象の選択と壁打ち")
    
    options_map = {"main": "メイン記事", "sub1": "記事2", "sub2": "記事3"}
    valid_options = [k for k, v in st.session_state.ocr_results.items() if len(v) > 10]
    
    if not valid_options:
        st.warning("OCRデータがありません。Tab 1で解析を行ってください。")
        selected_article_text = ""
    else:
        selected_key = st.radio("対象記事を選択", valid_options, format_func=lambda x: options_map[x], horizontal=True)
        selected_article_text = st.session_state.ocr_results[selected_key]
        
        # 選択変更の検知
        if selected_key != st.session_state.selected_article_key:
            st.session_state.selected_article_key = selected_key
            # 切替時にドラフトをクリアしたい場合は以下を有効化
            # st.session_state.current_draft = "" 
            # st.session_state.chat_history = []
            st.toast(f"{options_map[selected_key]} に切り替えました")

    st.markdown("---")

    col_draft, col_chat = st.columns([1, 1])

    with col_draft:
        st.markdown("### 📝 感想文ドラフト")
        
        if st.button("🚀 初稿を作成する", disabled=(not selected_article_text)):
            if not client:
                 st.error("OpenAI APIキーがありません。")
            else:
                with st.spinner("執筆中..."):
                    draft = generate_draft(selected_article_text, "", target_length)
                    st.session_state.current_draft = draft
                    st.session_state.chat_history = [{
                        "role": "assistant", 
                        "content": "初稿を作成しました！\nより良い感想文にするために、この記事に関連するあなたの具体的な体験談を教えてください。"
                    }]
                    st.rerun()
        
        if st.session_state.current_draft:
            st.text_area("現在の原稿", st.session_state.current_draft, height=600, key="draft_area")
            
            if st.button("🔄 チャット反映して書き直し", type="primary"):
                with st.spinner("リライト中..."):
                    chat_context = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.chat_history])
                    st.session_state.current_draft = generate_draft(selected_article_text, chat_context, target_length)
                    st.success("完了！")
                    st.rerun()

    with col_chat:
        st.markdown("### 💬 壁打ち")
        chat_container = st.container(height=500)
        
        for message in st.session_state.chat_history:
            with chat_container.chat_message(message["role"]):
                st.markdown(message["content"])

        if prompt := st.chat_input("エピソードを入力..."):
            if not selected_article_text:
                st.error("先に記事を選択して初稿を作成してください。")
            elif not client:
                st.error("OpenAI APIキーがありません。")
            else:
                st.session_state.chat_history.append({"role": "user", "content": prompt})
                with chat_container.chat_message("user"):
                    st.markdown(prompt)

                with chat_container.chat_message("assistant"):
                    with st.spinner("考え中..."):
                        chat_sys = f"あなたは編集者です。以下の記事内容を踏まえ、ユーザーから深いエピソードを引き出してください。\n記事: {selected_article_text[:500]}..."
                        msgs = [{"role": "system", "content": chat_sys}] + st.session_state.chat_history
                        res = client.chat.completions.create(model="gpt-4o", messages=msgs)
                        ai_res = res.choices[0].message.content
                        
                st.markdown(ai_res)
                st.session_state.chat_history.append({"role": "assistant", "content": ai_res})

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
    else:
        st.info("感想文を作成し、テンプレートをアップロードしてください。")
