import streamlit as st
import google.generativeai as genai
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
from PIL import Image
import concurrent.futures
import time
import random

# ==========================================
# ページ設定（必ず最初）
# ==========================================
st.set_page_config(
    page_title="致知読書感想文アプリ v5.1（タグOCR・安定UI版）",
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
""".strip()

# Excel書き込み設定
EXCEL_START_ROW = 9
CHARS_PER_LINE = 40
EXCEL_CLEAR_ROWS = 500  # ここまで消す（残骸対策）

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
# サイドバー：API・設定
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
        except Exception:
            st.error("OpenAIキーが無効です")

    if google_key:
        try:
            genai.configure(api_key=google_key)
        except Exception:
            st.error("Googleキーが無効です")

    st.markdown("---")
    uploaded_template = st.file_uploader("感想文フォーマット(.xlsx)", type=["xlsx"])
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)

    st.markdown("---")
    st.caption("🔧 OCRモデル/分割設定")
    model_main = st.text_input("メインModel ID", value="gemini-3-flash-preview")
    model_sub = st.text_input("サブModel ID", value="gemini-2.0-flash-lite-preview-02-05")

    # 段組の強制分割（一般的な誌面は3段が多い想定）
    col_splits = st.selectbox("段組（列）分割数", [2, 3, 4], index=1)
    row_splits = st.selectbox("上下分割数", [1, 2, 3], index=1)

    max_workers = st.selectbox("OCR並列数（429対策）", [1, 2, 3], index=1)

    st.markdown("---")
    if st.button("🗑️ リセット"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# ==========================================
# 関数
# ==========================================
def split_text(text: str, chunk_size: int):
    if not text:
        return []
    clean_text = text.replace("\n", "　")
    return [clean_text[i:i + chunk_size] for i in range(0, len(clean_text), chunk_size)]

def pil_from_uploads(uploaded_files):
    imgs = []
    if not uploaded_files:
        return imgs
    for f in uploaded_files:
        imgs.append(Image.open(f).convert("RGB"))
    return imgs

def crop_segments(img: Image.Image, cols: int = 3, rows: int = 2):
    """
    画像を cols×rows に分割して、縦書き誌面の読み順に並べる。
    読み順：右列→左列、各列は上→下（rowsが2なら上段→下段）
    返り値：[(label, segment_image), ...]
    """
    w, h = img.size
    col_w = w // cols
    row_h = h // rows

    segments = []
    # 右→左
    for c in range(cols):
        col_index_from_right = cols - 1 - c
        x0 = col_index_from_right * col_w
        x1 = w if col_index_from_right == cols - 1 else x0 + col_w  # 端は誤差吸収
        for r in range(rows):
            y0 = r * row_h
            y1 = h if r == rows - 1 else y0 + row_h
            seg = img.crop((x0, y0, x1, y1))
            # ラベル（例：右列上 / 中列下）
            col_name = ["左列", "中列", "右列"]
            # colsが2/4のときもそれっぽく命名
            if cols == 2:
                col_label = "右列" if col_index_from_right == 1 else "左列"
            elif cols == 3:
                col_label = col_name[col_index_from_right]
            else:
                # 4列以上は番号で
                col_label = f"{col_index_from_right+1}列目(右起点)"
            row_label = ["上", "中", "下"][r] if rows <= 3 else f"{r+1}段目"
            label = f"{col_label}{row_label}"
            segments.append((label, seg))
    return segments

def gemini_generate_with_retry(model_id: str, inputs, retries: int = 4):
    """
    Gemini呼び出しを指数バックオフでリトライ（429/一時エラー対策）
    """
    last_err = None
    for i in range(retries + 1):
        try:
            model = genai.GenerativeModel(model_id)
            res = model.generate_content(inputs)
            return res.text
        except Exception as e:
            last_err = e
            # バックオフ：1.0, 2.0, 4.0... + ジッタ
            sleep_s = (2 ** i) + random.uniform(0, 0.6)
            time.sleep(sleep_s)
    raise last_err

def process_ocr_tagged(label: str, uploaded_files, model_id: str, cols: int, rows: int):
    """
    引用用：タグ付きOCR（[ファイル名][セグメント]）
    """
    if not uploaded_files:
        return ""

    # 注意：uploaded_files の順序はユーザーの選択順になりがちだが、環境で変わる場合あり。
    # 安全にするなら file.name でソートも可。ここは“選んだ順”重視でそのまま。
    system_prompt = (
        "あなたは高精度OCRエンジンです。\n"
        "以下の雑誌『致知』画像から、書いてある文字を一字一句そのまま書き起こしてください。\n"
        "【厳守】要約・省略・言い換え禁止。判読不能は(判読不能)。\n"
        "縦書きは右段→左段の順で読む。段をまたいで1行として読まない。\n"
        "必ず位置タグを付ける：\n"
        "  [ファイル名: xxx]\n"
        "  <セグメント: 右列上> ...本文...\n"
        "のように出力する。\n"
        "※画像は1ページを cols×rows に分割したセグメントが、読み順で送られる。\n"
    )

    gemini_inputs = [system_prompt]

    for f in uploaded_files:
        # Streamlit UploadedFile: name 属性あり
        fname = getattr(f, "name", "unknown")
        img = Image.open(f).convert("RGB")
        segs = crop_segments(img, cols=cols, rows=rows)

        gemini_inputs.append(f"\n\n[ファイル名: {fname}]\n")
        for seg_label, seg_img in segs:
            gemini_inputs.append(f"<セグメント: {seg_label}>\n")
            gemini_inputs.append(seg_img)
            gemini_inputs.append("\n")  # 区切り

    try:
        text = gemini_generate_with_retry(model_id, gemini_inputs, retries=4)
        # 念のため、記事ラベルも先頭に付ける（後処理しやすい）
        return f"=== {label} ===\n{text}"
    except Exception as e:
        return f"=== {label} ===\n[エラー: OCR失敗: {e}]"

def generate_draft(article_text: str, chat_context: str, target_len: int):
    if not client:
        return "エラー: OpenAI APIキーが設定されていません。"

    system_prompt = (
        "あなたは税理士事務所の職員です。\n"
        "雑誌『致知』の読書感想文（社内木鶏会用）を作成します。\n"
        "【過去の感想文】を分析し、文体・熱量・業務への結びつけ方を模倣してください。"
    )

    user_content = (
        f"【今回選択した記事のOCRデータ】\n{article_text}\n\n"
        f"【ユーザーの過去の感想文（スタイル見本）】\n{PAST_REVIEWS}\n\n"
        f"【打ち合わせ内容】\n{chat_context}\n\n"
        "【執筆条件】\n"
        f"- 文字数：{target_len}文字前後\n"
        "- 文体：「です・ます」調\n"
        "- 段落ごとに改行を入れること\n"
        "- 構成：①記事の引用 ②自分の業務エピソード ③今後の決意\n"
    )

    res = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content},
        ],
        temperature=0.7
    )
    return res.choices[0].message.content

# ==========================================
# UI
# ==========================================
st.title("📖 致知読書感想文アプリ v5.1（タグOCR・安定UI版）")
st.caption("Step 1: タグ付きOCR → Step 2: 記事選択・執筆 → Step 3: Excel出力")

tab1, tab2, tab3 = st.tabs(["1️⃣ 画像解析（タグOCR）", "2️⃣ 記事選択 & 執筆", "3️⃣ Excel出力"])

# ------------------------------------------
# Tab 1: OCR
# ------------------------------------------
with tab1:
    st.subheader("Step 1. 記事画像の読み込み")
    st.info("段組混線を防ぐため、画像を『右→左』の列分割＋上下分割し、位置タグ付きでOCRします。")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("#### 📂 メイン記事")
        files_main = st.file_uploader("画像を選択", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="f1")
    with col2:
        st.markdown("#### 📂 記事2")
        files_sub1 = st.file_uploader("画像を選択", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="f2")
    with col3:
        st.markdown("#### 📂 記事3")
        files_sub2 = st.file_uploader("画像を選択", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="f3")

    if st.button("🚀 全記事を一括解析（並列）", type="primary"):
        if not (files_main or files_sub1 or files_sub2):
            st.error("画像が選択されていません。")
        elif not google_key:
            st.error("Google APIキーが設定されていません。")
        else:
            with st.spinner("タグ付きOCR中...（レート制限時は自動リトライ）"):
                try:
                    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
                        f_main = ex.submit(process_ocr_tagged, "メイン記事", files_main, model_main, col_splits, row_splits)
                        f_sub1 = ex.submit(process_ocr_tagged, "記事2", files_sub1, model_sub, col_splits, row_splits)
                        f_sub2 = ex.submit(process_ocr_tagged, "記事3", files_sub2, model_sub, col_splits, row_splits)

                        st.session_state.ocr_results["main"] = f_main.result()
                        st.session_state.ocr_results["sub1"] = f_sub1.result()
                        st.session_state.ocr_results["sub2"] = f_sub2.result()

                    st.success("✅ 解析完了！ '2️⃣ 記事選択 & 執筆' タブへ。")
                except Exception as e:
                    st.error(f"予期せぬエラー: {e}")

    with st.expander("OCR解析結果を確認する"):
        st.text_area("Main", st.session_state.ocr_results["main"], height=200)
        st.text_area("Sub1", st.session_state.ocr_results["sub1"], height=200)
        st.text_area("Sub2", st.session_state.ocr_results["sub2"], height=200)

# ------------------------------------------
# Tab 2: Draft & Chat
# ------------------------------------------
with tab2:
    st.subheader("Step 2. 執筆対象の選択と壁打ち")

    options_map = {"main": "メイン記事", "sub1": "記事2", "sub2": "記事3"}
    valid_options = [k for k, v in st.session_state.ocr_results.items() if len(v) > 20]

    if not valid_options:
        st.warning("OCRデータがありません。Tab 1で解析してください。")
        selected_article_text = ""
    else:
        selected_key = st.radio(
            "対象記事を選択",
            valid_options,
            format_func=lambda x: options_map[x],
            horizontal=True
        )
        selected_article_text = st.session_state.ocr_results[selected_key]

        if selected_key != st.session_state.selected_article_key:
            st.session_state.selected_article_key = selected_key
            st.toast(f"{options_map[selected_key]} に切り替えました")

    st.markdown("---")

    col_draft, col_chat = st.columns([1, 1])

    # --- Draft column
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
                        "content": "初稿を作成しました！この記事に関連するあなたの具体的な体験談を教えてください。"
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

    # --- Chat column
    with col_chat:
        st.markdown("### 💬 壁打ち（右カラム内で安定表示）")

        chat_box = st.container(height=420)
        for m in st.session_state.chat_history:
            with chat_box.chat_message(m["role"]):
                st.markdown(m["content"])

        # chat_input は位置が不安定になりやすいので text_input + button に変更
        st.markdown("#### 入力")
        user_msg = st.text_input("エピソードを入力…", key="chat_text_input")
        send = st.button("送信", type="secondary")

        if send and user_msg.strip():
            if not selected_article_text:
                st.error("先に記事を選択して初稿を作成してください。")
            elif not client:
                st.error("OpenAI APIキーがありません。")
            else:
                st.session_state.chat_history.append({"role": "user", "content": user_msg.strip()})

                # 編集者として深掘り質問を作る
                with st.spinner("考え中..."):
                    chat_sys = (
                        "あなたは編集者です。ユーザーから『具体的な体験談』を引き出すために、"
                        "深掘り質問を1〜3個、日本語で作ってください。\n"
                        "記事内容（先頭一部）:\n"
                        f"{selected_article_text[:800]}"
                    )
                    msgs = [{"role": "system", "content": chat_sys}] + st.session_state.chat_history[-8:]
                    res = client.chat.completions.create(model="gpt-4o", messages=msgs, temperature=0.7)
                    ai_res = res.choices[0].message.content

                st.session_state.chat_history.append({"role": "assistant", "content": ai_res})

                # 入力欄クリア
                st.session_state.chat_text_input = ""
                st.rerun()

# ------------------------------------------
# Tab 3: Excel output
# ------------------------------------------
with tab3:
    st.subheader("Step 3. Excel出力")

    if st.session_state.current_draft and uploaded_template:
        if st.button("📥 Excelダウンロード"):
            try:
                wb = load_workbook(uploaded_template)
                ws = wb.active

                # クリア（残骸防止）
                for r in range(EXCEL_START_ROW, EXCEL_START_ROW + EXCEL_CLEAR_ROWS):
                    ws[f"A{r}"].value = None

                lines = split_text(st.session_state.current_draft, CHARS_PER_LINE)
                for i, line in enumerate(lines):
                    cell = ws[f"A{EXCEL_START_ROW + i}"]
                    cell.value = line
                    cell.alignment = Alignment(wrap_text=False, shrink_to_fit=False, horizontal="left")

                out = io.BytesIO()
                wb.save(out)
                out.seek(0)

                st.download_button(
                    "Excel保存",
                    out,
                    "感想文.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("完了！")
            except Exception as e:
                st.error(f"エラー: {e}")
    else:
        st.info("感想文を作成し、テンプレートをアップロードしてください。")
