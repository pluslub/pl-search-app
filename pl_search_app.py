import os
import io
import msal
import requests
import streamlit as st
import google.generativeai as genai
from docx import Document
import openpyxl

# --- 設定情報（Streamlit Secretsから取得）---
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
MS_CLIENT_ID = st.secrets["MS_CLIENT_ID"]
MS_TENANT_ID = st.secrets["MS_TENANT_ID"]

genai.configure(api_key=GEMINI_API_KEY)

# --- セッション状態の初期化 ---
if "ms_token" not in st.session_state:
    st.session_state.ms_token = None
if "search_results" not in st.session_state:
    st.session_state.search_results = []
if "selected_items" not in st.session_state:
    st.session_state.selected_items = []
if "analysis_result" not in st.session_state:
    st.session_state.analysis_result = ""
if "device_flow" not in st.session_state:
    st.session_state.device_flow = None
if "msal_app" not in st.session_state:
    st.session_state.msal_app = None

# --- MSALアプリの初期化 ---
def get_msal_app():
    if st.session_state.msal_app is None:
        authority = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
        st.session_state.msal_app = msal.PublicClientApplication(
            MS_CLIENT_ID,
            authority=authority
        )
    return st.session_state.msal_app

# --- AIモデル自動選択 ---
def get_working_model():
    try:
        available_models = [
            m.name for m in genai.list_models()
            if 'generateContent' in m.supported_generation_methods
        ]
        if not available_models:
            raise Exception("利用可能なモデルが見つかりません。")
        target = next((m for m in available_models if "1.5-flash" in m), None)
        if not target:
            target = next((m for m in available_models if "flash" in m), None)
        if not target:
            target = available_models[0]
        return genai.GenerativeModel(target)
    except Exception as e:
        return genai.GenerativeModel("gemini-1.5-flash")

# --- OneDriveファイルのテキスト抽出 ---
def extract_text_from_ms_file(item_id, file_name, token):
    content_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{item_id}/content"
    headers = {'Authorization': f'Bearer {token}'}
    res = requests.get(content_url, headers=headers)

    if res.status_code != 200:
        return ""

    file_bytes = io.BytesIO(res.content)
    text = ""

    try:
        if file_name.endswith(('.xlsx', '.xlsm')):
            wb = openpyxl.load_workbook(file_bytes, data_only=True)
            for sheet in wb.worksheets:
                text += f"\n[シート: {sheet.title}]\n"
                for row in sheet.iter_rows(values_only=True):
                    row_data = " ".join([str(cell) for cell in row if cell is not None])
                    if row_data.strip():
                        text += row_data + "\n"
        elif file_name.endswith('.docx'):
            doc = Document(file_bytes)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif file_name.endswith('.txt'):
            text = res.content.decode('utf-8', errors='ignore')
    except Exception as e:
        st.warning(f"解析エラー({file_name}): {e}")

    return text[:4000]

# ======================
# UI
# ======================
st.title("📂 OneDrive ファイル検索アプリ")

# --- Microsoft認証 ---
st.header("① Microsoft 365 にログイン")

app = get_msal_app()
accounts = app.get_accounts()

if st.session_state.ms_token:
    st.success("✅ ログイン済みです")
    if st.button("ログアウト"):
        st.session_state.ms_token = None
        st.session_state.device_flow = None
        st.rerun()

elif accounts:
    result = app.acquire_token_silent(["Files.Read.All"], account=accounts[0])
    if result and "access_token" in result:
        st.session_state.ms_token = result["access_token"]
        st.success("✅ ログイン済みです")
        st.rerun()

else:
    if st.session_state.device_flow is None:
        if st.button("Microsoft 365 でログイン"):
            flow = app.initiate_device_flow(scopes=["Files.Read.All"])
            st.session_state.device_flow = flow
            st.rerun()
    else:
        flow = st.session_state.device_flow
        st.info(f"以下のURLにアクセスしてコードを入力してください：\n\n**{flow['verification_uri']}**\n\nコード：**{flow['user_code']}**")
        if st.button("ログイン完了（認証後にクリック）"):
            with st.spinner("認証確認中..."):
                result = app.acquire_token_by_device_flow(flow)
                if result and "access_token" in result:
                    st.session_state.ms_token = result["access_token"]
                    st.session_state.device_flow = None
                    st.success("✅ ログイン成功！")
                    st.rerun()
                else:
                    st.error("❌ 認証に失敗しました。もう一度お試しください。")
                    st.session_state.device_flow = None

# --- OneDrive検索 ---
if st.session_state.ms_token:
    st.header("② OneDrive を検索")
    keyword = st.text_input("検索キーワードを入力してください")

    if st.button("🔍 検索"):
        if not keyword:
            st.warning("キーワードを入力してください。")
        else:
            with st.spinner("OneDriveを検索中..."):
                headers = {'Authorization': f'Bearer {st.session_state.ms_token}'}
                url = f"https://graph.microsoft.com/v1.0/me/drive/root/search(q='{keyword}')"
                response = requests.get(url, headers=headers)

                if response.status_code != 200:
                    st.error(f"検索エラー: {response.status_code}")
                else:
                    items = response.json().get('value', [])
                    if not items:
                        st.warning("一致するファイルが見つかりませんでした。")
                    else:
                        st.session_state.search_results = items
                        st.session_state.selected_items = []
                        st.session_state.analysis_result = ""

    # --- 検索結果表示 ---
    if st.session_state.search_results:
        st.header("③ 解析するファイルを選択")
        st.write(f"{len(st.session_state.search_results)} 件見つかりました")

        selected = []
        for i, item in enumerate(st.session_state.search_results):
            name = item['name']
            url = item.get('webUrl', '')
            checked = st.checkbox(f"{name}", key=f"file_{i}")
            if url:
                st.markdown(f"　🔗 [{url}]({url})", unsafe_allow_html=True)
            if checked:
                selected.append(item)

        if st.button("🤖 選択したファイルをAIで解析"):
            if not selected:
                st.warning("ファイルを1つ以上選択してください。")
            else:
                with st.spinner("ファイルを読み取り中..."):
                    analysis_context = ""
                    for item in selected:
                        name = item['name']
                        st.write(f"読み取り中: {name}")
                        content_text = extract_text_from_ms_file(
                            item['id'], name, st.session_state.ms_token
                        )
                        if content_text:
                            analysis_context += f"\n--- ファイル名: {name} ---\n{content_text}\n"

                if not analysis_context:
                    st.warning("解析可能なテキストが見つかりませんでした。")
                else:
                    with st.spinner("AIが分析中..."):
                        model = get_working_model()
                        keyword_used = st.session_state.get("last_keyword", "")
                        prompt = (
                            f"あなたは優秀なアシスタントです。以下のファイルから"
                            f"関連する具体的な数値、日付、指示、変更点を抜き出して要約してください。\n\n"
                            f"【解析対象データ】\n{analysis_context}"
                        )
                        try:
                            ai_res = model.generate_content(prompt)
                            st.session_state.analysis_result = ai_res.text.strip()
                        except Exception as e:
                            st.error(f"AI分析エラー: {e}")

    # --- 分析結果表示 ---
    if st.session_state.analysis_result:
        st.header("④ AIの分析結果")
        st.markdown(st.session_state.analysis_result)
