import streamlit as st
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import base64

# ==========================================
# ページ設定・定数定義
# ==========================================
st.set_page_config(page_title="致知読書感想文アプリ", layout="wide", page_icon="📖")
st.title("📖 致知読書感想文作成アプリ")
st.caption("Step 1：記事の解析（事実抽出） → Step 2：感想文執筆（税理士事務所向け）")

# Excel書き込み設定
EXCEL_START_ROW = 9  # 書き込み開始行（A9）
CHARS_PER_LINE = 40  # 1行あたりの文字数

# ==========================================
# 関数定義
# ==========================================

def get_openai_client():
    """OpenAIクライアントの初期化"""
    api_key = st.secrets.get("OPENAI_API_KEY")
    if not api_key:
        st.error("⚠️ OpenAI APIキーが設定されていません。.streamlit/secrets.toml を確認してください。")
        st.stop()
    return OpenAI(api_key=api_key)

def encode_image(image_file):
    """画像をBase64文字列に変換"""
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

def split_text(text, chunk_size):
    """テキストを指定文字数ごとに分割するジェネレータ"""
    # Excelの行に合わせて改行コードを除去して詰める場合と、
    # 段落を維持する場合があるが、今回は「指定フォーマットへ流し込み」のため
    # 連続した文字列として扱い、単純分割を行う
    clean_text = text.replace('\n', '　')  # 改行を全角スペースに置換して行ズレを防ぐ
    return [clean_text[i:i+chunk_size] for i in range(0, len(clean_text), chunk_size)]

# ==========================================
# セッション状態の管理
# ==========================================
if "extracted_text" not in st.session_state:
    st.session_state.extracted_text = ""
if "final_text" not in st.session_state:
    st.session_state.final_text = ""

client = get_openai_client()

# ==========================================
# サイドバー（設定）
# ==========================================
with st.sidebar:
    st.header("⚙️ 設定・テンプレート")
    uploaded_template = st.file_uploader(
        "感想文フォーマット(.xlsx)", 
        type=["xlsx"],
        help="A9セルから書き込みが開始されます"
    )
    target_length = st.selectbox("目標文字数", [300, 400, 500, 600, 700, 800], index=1)

# ==========================================
# Step 1: 記事画像の読み込みと解析
# ==========================================
st.header("Step 1. 記事画像の解析")
st.info("💡 15枚以上の画像も一括でアップロード可能です。ファイル名順（IMG_001...）に処理されます。")

uploaded_files = st.file_uploader(
    "画像をまとめて選択（ドラッグ＆ドロップ可）", 
    type=['png', 'jpg', 'jpeg'], 
    accept_multiple_files=True
)

if uploaded_files:
    # ファイル数表示
    st.write(f"📁 {len(uploaded_files)}枚の画像を読み込みました")

    if st.button("🔍 全ページを解析して引用元を抽出する", type="primary"):
        with st.spinner("AIが画像を順に解析中...（これには少し時間がかかります）"):
            try:
                # 1. 画像順序の保証（ファイル名でソート）
                uploaded_files.sort(key=lambda x: x.name)
                
                content_list = []
                
                # 2. システムプロンプト（事実抽出・創造性排除）
                system_prompt = """
                あなたは雑誌『致知』の記事解析専門AIです。以下の厳格なルールに従い、画像から情報を抽出してください。

                【役割】
                提供された全ページの画像（順番通り）を読み込み、感想文の根拠となる「事実」と「引用」のみを抽出する。
                
                【絶対遵守事項】
                1. 温度感（Temperature）は0として扱い、記事に書かれていないことは一切創作しない。
                2. 重要な文章を抜き出す際は、**必ず「掲載位置」を正確に付記**すること。
                   例：「〜である」（1枚目 右段 5行目付近）
                   例：「〜という言葉が胸に響く」（3枚目 左段 写真キャプション）
                3. 記事全体の要約を作成すること。
                4. 登場人物の名前、肩書きは正確に転記すること。
                """
                content_list.append({"type": "text", "text": system_prompt})

                # 3. 画像データの構築
                for i, img_file in enumerate(uploaded_files):
                    base64_img = encode_image(img_file)
                    # AIにページ番号とファイル名を明示
                    content_list.append({
                        "type": "text", 
                        "text": f"\n【画像 {i+1}枚目 (ファイル名: {img_file.name})】\n"
                    })
                    content_list.append({
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}
                    })

                # 4. APIコール（事実抽出用）
                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": content_list}],
                    max_tokens=4000,
                    temperature=0.0  # 創造性ゼロ
                )

                st.session_state.extracted_text = response.choices[0].message.content
                st.session_state.final_text = "" # 再解析したら感想文もクリア
                st.rerun()

            except Exception as e:
                st.error(f"解析エラーが発生しました: {e}")

# ==========================================
# 解析結果の編集エリア
# ==========================================
if st.session_state.extracted_text:
    st.markdown("---")
    st.subheader("📝 解析結果（事実確認）")
    st.caption("AIが抽出したテキストです。間違いがある場合や、補足したい場合はここで修正してください。")
    
    # ユーザーが編集可能
    edited_text = st.text_area(
        "抽出されたテキスト（ここを編集するとStep 2に反映されます）", 
        st.session_state.extracted_text, 
        height=400
    )
    st.session_state.extracted_text = edited_text

    # ==========================================
    # Step 2: 感想文の作成
    # ==========================================
    st.markdown("---")
    st.header("Step 2. 感想文の執筆")

    if st.button("✍️ 税理士事務所員として感想文を書く"):
        with st.spinner("執筆中..."):
            try:
                # 執筆用プロンプト
                writer_prompt = f"""
                あなたは税理士事務所の誠実な職員です。
                社内木鶏会のために、以下の【解析済み記事データ】をもとに読書感想文を作成してください。

                【解析済み記事データ】
                {st.session_state.extracted_text}

                【構成ルール】
                1. 記事の要約（簡潔に）
                2. 印象に残った言葉（※解析データの引用元記述を活用し、正確に引用すること）
                3. 自分の業務（税理士補助業務・顧客対応など）への具体的な活かし方

                【執筆条件】
                - 文字数：{target_length}文字前後
                - 文体：「です・ます」調
                - タイトルは不要。本文のみを出力。
                - 段落ごとに改行を入れること。
                - 決して嘘や記事にないエピソードをでっち上げないこと。
                """

                res = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": writer_prompt}],
                    temperature=0.7  # 執筆時は多少の表現の幅を持たせる
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
    st.subheader("🎉 完成プレビュー & 出力")
    
    # プレビュー表示
    st.text_area("完成した感想文", st.session_state.final_text, height=300)

    if uploaded_template:
        try:
            # メモリ上でExcel操作
            wb = load_workbook(uploaded_template)
            ws = wb.active

            # 既存の入力内容をクリア（A9以降）
            for row in range(EXCEL_START_ROW, 100):
                cell = ws[f"A{row}"]
                cell.value = None

            # 40文字ずつ分割してリスト化
            lines = split_text(st.session_state.final_text, CHARS_PER_LINE)

            # 書き込み処理
            for i, line in enumerate(lines):
                target_row = EXCEL_START_ROW + i
                cell = ws[f"A{target_row}"]
                cell.value = line
                # 縮小して全体を表示 & 折り返さない
                cell.alignment = Alignment(shrink_to_fit=True, wrap_text=False)

            # 保存用バッファ
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)

            st.download_button(
                label="📥 Excelファイルを出力する",
                data=out,
                file_name="致知感想文_完成.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:
            st.error(f"Excel出力エラー: {e}")
    else:
        st.warning("⚠️ Excelテンプレートがアップロードされていません。サイドバーからアップロードしてください。")
