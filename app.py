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
    page_title="致知読書感想文アプリ v5.2",
    layout="wide",
    page_icon="📖"
)

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
    【OCR関数】レイアウト認識強化版
    """
    if not pil_images:
        return ""
    
    try:
        gemini_inputs = []
        system_prompt = (
            "あなたは高精度なOCRエンジンです。雑誌『致知』の紙面を読み取ります。\n"
            "【重要ルール】\n"
            "1. 画像全体を見て、レイアウト（段組み）を認識してください。\n"
            "2. 記事のブロック（意味のまとまり）ごとに読み進めてください。\n"
            "3. 縦書きの段組みがある場合、右の段から左の段へと順番に読み、段をまたいで一行として読まないように注意してください。\n"
            "4. 複数の記事がある場合は、記事ごとに区切って出力してください。\n"
            "5. 出力形式: [画像番号] <本文>..."
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
    """
    【修正版】感想文生成関数
    chat_context（壁打ち内容）の有無によって、プロンプトを完全に切り替える。
    これにより「初稿での妄想」を防ぎ、「書き直し」での確実な反映を実現する。
    """
    if not client:
        return "エラー: OpenAI APIキーが設定されていません。"

    # ====================================================
    # パターンA: 初稿作成（チャットなし）
    # ====================================================
    if not chat_context:
        system_prompt = (
            "あなたは税理士事務所の職員です。\n"
            "雑誌『致知』の読書感想文（社内木鶏会用）の【初稿】を作成します。\n"
            "以下の【ユーザーの過去の感想文】の文体や熱量を模倣し、"
            "【記事の内容】をベースに感想文を書いてください。\n"
            "**重要: まだ具体的なエピソードは入力されていないため、勝手に具体的な体験談を創作しないでください。**\n"
            "業務への結びつけは「日々の業務において〜」といった一般的な表現に留めてください。"
        )
        user_content = (
            f"【今回選択した記事のOCRデータ】\n{article_text}\n\n"
            f"【ユーザーの過去の感想文（文体見本）】\n{PAST_REVIEWS}\n\n"
            "【執筆条件】\n"
            f"- 文字数：{target_len}文字前後\n"
            "- 文体：「です・ます」調\n"
            "- 段落ごとに改行を入れること。\n"
            "- 構成：①記事の引用・要約 ②一般的な業務への気づき（※創作エピソード禁止） ③今後の決意"
        )

    # ====================================================
    # パターンB: 書き直し（壁打ち反映）
    # ====================================================
    else:
        system_prompt = (
            "あなたは税理士事務所の職員です。\n"
            "読書感想文の【書き直し】を行います。\n"
            "これまでの感想文（または記事内容）に、**【壁打ちチャットでの追加エピソード】を具体的に組み込んでください。**\n"
            "抽象的だった「業務への気づき」の部分を、チャットで語られた「具体的な体験談」に完全に差し替えてください。"
        )
        user_content = (
            f"【記事データ】\n{article_text}\n\n"
            f"【ユーザーの過去の感想文（文体見本）】\n{PAST_REVIEWS}\n\n"
            f"【壁打ちチャットでの追加エピソード（※これを必ず書くこと※）】\n{chat_context}\n\n"
            "【執筆条件】\n"
            f"- 文字数：{target_len}文字前後\n"
            "- 文体：「です・ます」調\n"
            "- 構成：①記事の引用 ②**チャットで出た具体的なエピソード** ③今後の決意\n"
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
st.title("📖 致知読書感想文アプリ v5.2 (初稿/リライト分離版)")
st.caption("Step 1: 全体レイアウト解析OCR → Step 2: 記事選択・執筆 → Step 3: Excel出力")

tab1, tab2, tab3 = st.tabs(["1️⃣ 画像解析", "2️⃣ 記事選択 & 執筆", "3️⃣ Excel出力"])

# ------------------------------------------------------------------
# Tab 1: 並列OCR処理
# ------------------------------------------------------------------
with tab1:
    st.subheader("Step 1. 記事画像の読み込み")
    st.info("画像を分割せず、AIにレイアウト全体を認識させることで正確に読み取ります。")

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
            with st.spinner("レイアウトを解析して読み取っています..."):
                try:
                    images_main = [Image.open(f).convert("RGB") for f in files_main] if files_main else []
                    images_sub1 = [Image.open(f).convert("RGB") for f in files_sub1] if files_sub1 else []
                    images_sub2 = [Image.open(f).convert("RGB") for f in files_sub2] if files_sub2 else []

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
        
        if selected_key != st.session_state.selected_article_key:
            st.session_state.selected_article_key = selected_key
            st.toast(f"{options_map[selected_key]} に切り替えました")

    st.markdown("---")

    col_draft, col_chat = st.columns([1, 1])

    with col_draft:
        st.markdown("### 📝 感想文ドラフト")
        
        # 初稿作成ボタン
        if st.button("🚀 初稿を作成する", disabled=(not selected_article_text)):
            if not client:
                 st.error("OpenAI APIキーがありません。")
            else:
                with st.spinner("初稿を作成中（エピソードはまだ入れません）..."):
                    # チャット履歴をリセット
                    st.session_state.chat_history = [] 
                    # 第2引数をNoneにすることで「初稿モード」にする
                    draft = generate_draft(selected_article_text, None, target_length)
                    st.session_state.current_draft = draft
                    
                    # AIからの最初の質問を履歴に追加
                    st.session_state.chat_history.append({
                        "role": "assistant", 
                        "content": "初稿を作成しました！\n今の段階では一般的な内容になっています。\n\n**この記事のテーマに関連して、あなたの業務で起きた具体的な出来事（成功・失敗・気づき）を教えてください。感想文に反映させます。**"
                    })
                    st.rerun()
        
        if st.session_state.current_draft:
            st.text_area("現在の原稿", st.session_state.current_draft, height=600, key="draft_area")
            
            # 書き直しボタン
            if st.button("🔄 チャットの内容を反映して書き直す", type="primary"):
                if not st.session_state.chat_history:
                    st.warning("まだチャットで会話していません。右側でエピソードを話してください。")
                else:
                    with st.spinner("チャットのエピソードを組み込んでリライト中..."):
                        # チャット履歴を文字列化
                        chat_context = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.chat_history])
                        
                        # 第2引数にチャット内容を渡すことで「書き直しモード」にする
                        new_draft = generate_draft(selected_article_text, chat_context, target_length)
                        st.session_state.current_draft = new_draft
                        st.success("書き直しました！")
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
