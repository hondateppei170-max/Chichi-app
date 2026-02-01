import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
from PIL import Image

# ==========================================
# 【重要】過去の感想文データ（文体学習用）
# ここにあなたの過去の感想文をコピペしてください。
# AIはこの文章の「書き出し」「熱量」「言葉選び」を真似します。
# ==========================================
PAST_REVIEWS = """
（例：以前の感想文1）
今月の致知を読んで、特に「逆境こそが人を育てる」という言葉が胸に刺さりました。
日々の税理士補助業務において、繁忙期にはつい愚痴が出そうになりますが、
それは自分の魂を磨く砥石なのだと気づかされました。
お客様の試算表を作る作業一つとっても、そこに魂を込めること。
それがプロフェッショナルとしての流儀だと感じます。

（例：以前の感想文2）
「一隅を照らす」という教えに、ハッとさせられました。
私はまだ大きな仕事は任されていませんが、コピー取りや掃除といった
足元の業務をおろそかにしていては、信頼は築けないと痛感しました。
明日からの巡回監査では、まず元気な挨拶から実践し、
お客様に元気をお届けできる存在になりたいと強く思います。
"""

# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="致知読書感想文アプリ v2", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ (壁打ち機能付き)")
st.caption("OCR(Gemini 3) → 執筆(GPT-4o) → 壁打ちで深掘り → Excel出力")

# Excel書き込み設定（A9セルから40文字ずつ）
EXCEL_START_ROW = 9
CHARS_PER_LINE = 40

# ==========================================
# API設定
# ==========================================
try:
    openai_key = st.secrets.get("OPENAI_API_KEY")
    if not openai_key:
        st.warning("⚠️ OpenAI APIキーが設定されていません。")
    else:
        client = OpenAI(api_key=openai_key)

    google_key = st.secrets.get("GOOGLE_API_KEY")
    if not google_key:
        st.warning("⚠️ Google APIキーが設定されていません。")
    else:
        genai.configure(api_key=google_key)
    
except Exception as e:
    st.error(f"API設定エラー: {e}")
    st.stop()

# ==========================================
# セッション状態の初期化
# ==========================================
if "extracted_text" not in st.session_state:
    st.session_state.extracted_text = ""  # OCR結果
if "current_draft" not in st.session_state:
    st.session_state.current_draft = ""   # 現在の感想文ドラフト
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []    # 壁打ちチャット履歴

# ==========================================
# 関数定義
# ==========================================
def split_text(text, chunk_size):
    """Excel用にテキストを指定文字数で分割"""
    if not text:
        return []
    clean_text = text.replace('\n', '　') # 改行は全角スペースに置換
    return [clean_text[i:i+chunk_size] for i in range(0, len(clean_text), chunk_size)]

def generate_draft(ocr_text, chat_context, target_len):
    """感想文を生成する関数（初回＆リライト共通）"""
    
    # 過去の文体 + チャットでの追加情報をプロンプトに組み込む
    system_prompt = (
        "あなたは税理士事務所の職員です。\n"
        "これから雑誌『致知』の読書感想文（社内木鶏会用）を作成します。\n"
        "以下の【ユーザーの過去の感想文】を徹底的に分析し、"
        "**「文体」「書き出しの癖」「精神的な熱量」「業務（巡回監査・決算など）への結びつけ方」**を模倣してください。\n"
        "単なる記事の要約ではなく、書き手の「体験」や「決意」が滲み出るような文章にしてください。"
    )

    user_content = (
        f"【OCR解析データ（記事内容）】\n{ocr_text}\n\n"
        f"【ユーザーの過去の感想文（スタイル見本）】\n{PAST_REVIEWS}\n\n"
        f"【これまでのチャットでの打ち合わせ内容（ここでのエピソードを必ず盛り込むこと）】\n{chat_context}\n\n"
        "【執筆条件】\n"
        f"- 文字数：{target_len}文字前後\n"
        "- 文体：「です・ます」調\n"
        "- 段落ごとに改行を入れること。\n"
        "- 構成：①記事で響いた言葉の引用（位置情報付き） ②そこから想起した自分の業務上のエピソード ③今後の決意"
    )

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content}
        ],
        temperature=0.7
    )
    return response.choices[0].message.content

# ==========================================
# サイドバー
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定")
    uploaded_template = st.file_uploader("感想文フォーマット(.xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)
    st.markdown("---")
    model_id_input = st.text_input("GeminiモデルID", value="gemini-3-flash-preview")
    
    # リセットボタン
    if st.button("🗑️ 全データをリセット"):
        for key in st.session_state.keys():
            del st.session_state[key]
        st.rerun()

# ==========================================
# メイン画面構成
# ==========================================
tab1, tab2, tab3 = st.tabs(["1️⃣ 画像解析 (OCR)", "2️⃣ 執筆 & 壁打ち (Chat)", "3️⃣ Excel出力"])

# ------------------------------------------------------------------
# Tab 1: OCR処理 (Gemini 3 Flash)
# ------------------------------------------------------------------
with tab1:
    st.subheader("Step 1. 記事画像の読み込み")
    st.info("Gemini 3 Flash を使用して、縦書き・段組みを正確に読み取ります。")
    
    col1, col2 = st.columns(2)
    with col1:
        files_main = st.file_uploader("メイン記事", type=['png', 'jpg', 'jpeg', 'webp'], accept_multiple_files=True)
    with col2:
        files_sub = st.file_uploader("サブ記事（任意）", type=['png', 'jpg', 'jpeg', 'webp'], accept_multiple_files=True)

    if st.button("🔍 解析開始", type="primary"):
        if not files_main:
            st.error("メイン記事の画像を選択してください。")
        else:
            with st.spinner(f"Gemini ({model_id_input}) が解析中..."):
                try:
                    gemini_inputs = []
                    # プロンプト（読み順の厳格な指定）
                    system_prompt_text = (
                        "あなたは、雑誌『致知』の紙面を完璧に読み取る高精度OCRエンジンです。\n"
                        "提供された画像は「上下2段組み」です。以下の順序を厳守してください。\n\n"
                        "1. 画像を上半分（上段）と下半分（下段）に分ける。\n"
                        "2. まず【上段】を右から左へ読む。\n"
                        "3. 次に【下段】を右から左へ読む。\n"
                        "※左段を上から下へ一気に読まないこと。\n\n"
                        "出力は [画像N枚目] <上段>... <下段>... のタグを付けてください。"
                    )
                    gemini_inputs.append(system_prompt_text)
                    
                    # 画像処理
                    all_files = files_main + (files_sub if files_sub else [])
                    for img_file in all_files:
                        img_file.seek(0)
                        image = Image.open(img_file).convert("RGB")
                        gemini_inputs.append(image)

                    # Gemini呼び出し
                    model = genai.GenerativeModel(model_id_input)
                    response = model.generate_content(gemini_inputs)
                    
                    st.session_state.extracted_text = response.text
                    st.success("✅ 解析完了！ '2️⃣ 執筆 & 壁打ち' タブへ移動してください。")
                
                except Exception as e:
                    st.error(f"OCRエラー: {e}")

    # OCR結果の確認・編集
    if st.session_state.extracted_text:
        st.markdown("---")
        st.session_state.extracted_text = st.text_area(
            "OCR結果（必要に応じて修正してください）", 
            st.session_state.extracted_text, 
            height=300
        )

# ------------------------------------------------------------------
# Tab 2: 執筆 & 壁打ち (Core Feature)
# ------------------------------------------------------------------
with tab2:
    st.subheader("Step 2. 感想文の執筆とブラッシュアップ")
    
    # レイアウト: 左側が感想文ドラフト、右側が壁打ちチャット
    col_draft, col_chat = st.columns([1, 1])

    # --- 左側：感想文表示エリア ---
    with col_draft:
        st.markdown("### 📝 感想文ドラフト")
        
        # 初回生成ボタン
        if not st.session_state.current_draft:
            if st.button("🚀 初稿を作成する"):
                if not st.session_state.extracted_text:
                    st.error("先にタブ1でOCRを行ってください。")
                else:
                    with st.spinner("過去の文体を分析して執筆中..."):
                        draft = generate_draft(st.session_state.extracted_text, "", target_length)
                        st.session_state.current_draft = draft
                        # 初回生成時にチャットのきっかけメッセージを入れる
                        st.session_state.chat_history.append({
                            "role": "assistant", 
                            "content": "初稿を作成しました！\nよりあなたらしい感想文にするために、少し質問させてください。\n\nこの記事の中で、ご自身の最近の業務（失敗談や成功体験）と重なる部分はありましたか？"
                        })
                        st.rerun()
        
        # ドラフトがある場合、表示
        if st.session_state.current_draft:
            st.text_area("現在の原稿", st.session_state.current_draft, height=600, key="draft_area")
            
            st.info("👈 右側のチャットでエピソードを追加し、下のボタンで書き直せます。")
            if st.button("🔄 チャットの内容を反映して書き直す", type="primary"):
                with st.spinner("会話内容を反映してリライト中..."):
                    # チャット履歴をテキスト化
                    chat_context = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.chat_history])
                    new_draft = generate_draft(st.session_state.extracted_text, chat_context, target_length)
                    st.session_state.current_draft = new_draft
                    st.success("書き直しました！")
                    st.rerun()

    # --- 右側：壁打ちチャットエリア ---
    with col_chat:
        st.markdown("### 💬 壁打ち (思考の深掘り)")
        st.caption("AIがあなたの専属編集者として質問します。答えることで、感想文に独自性が生まれます。")

        # チャットコンテナ
        chat_container = st.container(height=500)
        
        # 履歴の表示
        for message in st.session_state.chat_history:
            with chat_container.chat_message(message["role"]):
                st.markdown(message["content"])

        # ユーザー入力
        if prompt := st.chat_input("エピソードや考えを入力..."):
            # ユーザーのメッセージを追加
            st.session_state.chat_history.append({"role": "user", "content": prompt})
            with chat_container.chat_message("user"):
                st.markdown(prompt)

            # AIの返答（編集者モード）
            with chat_container.chat_message("assistant"):
                with st.spinner("考え中..."):
                    # チャット用プロンプト
                    chat_system = (
                        "あなたは、ユーザーの読書感想文をより良いものにするための『専属編集者』です。\n"
                        "ユーザーの発言を受け止め、さらに深いエピソード（具体的な業務での失敗、喜び、気づきなど）を引き出す質問を投げかけてください。\n"
                        "決して感想文そのものをここでは書かず、あくまでインタビューに徹してください。"
                    )
                    chat_messages = [{"role": "system", "content": chat_system}] + \
                                    [{"role": m["role"], "content": m["content"]} for m in st.session_state.chat_history]
                    
                    res = client.chat.completions.create(
                        model="gpt-4o",
                        messages=chat_messages,
                        temperature=0.7
                    )
                    ai_response = res.choices[0].message.content
                    
            st.markdown(ai_response)
            st.session_state.chat_history.append({"role": "assistant", "content": ai_response})

# ------------------------------------------------------------------
# Tab 3: Excel出力
# ------------------------------------------------------------------
with tab3:
    st.subheader("Step 3. Excelへの書き出し")
    
    if not st.session_state.current_draft:
        st.warning("まだ感想文が作成されていません。")
    else:
        st.write("完成した以下のテキストをExcelに出力します。")
        st.text(st.session_state.current_draft)
        
        if uploaded_template:
            if st.button("📥 Excelを作成してダウンロード"):
                try:
                    wb = load_workbook(uploaded_template)
                    ws = wb.active # アクティブシート（1枚目）を対象
                    
                    # 以前の内容をクリア（A9以降）
                    for row in range(EXCEL_START_ROW, 100):
                        ws[f"A{row}"].value = None
                    
                    # 40文字区切りでリスト化
                    lines = split_text(st.session_state.current_draft, CHARS_PER_LINE)
                    
                    # A9セルから順に書き込み
                    for i, line in enumerate(lines):
                        current_row = EXCEL_START_ROW + i
                        cell = ws[f"A{current_row}"]
                        cell.value = line
                        # 書式設定（折り返さない、縮小しない、左寄せ）
                        cell.alignment = Alignment(wrap_text=False, shrink_to_fit=False, horizontal='left')
                    
                    # バッファに保存
                    out = io.BytesIO()
                    wb.save(out)
                    out.seek(0)
                    
                    st.download_button(
                        label="Excelファイルを保存",
                        data=out,
                        file_name="社内木鶏会感想文.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("出力完了！")
                    
                except Exception as e:
                    st.error(f"Excel出力エラー: {e}")
        else:
            st.warning("テンプレートExcel（感想文フォーマット.xlsx）をサイドバーからアップロードしてください。")
