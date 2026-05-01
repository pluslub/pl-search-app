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

SCOPES = [
    "Files.Read.All",
    "Chat.Read",
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "ChannelMessage.Read.All",
]

# --- セッション状態の初期化 ---
for key in ["ms_token", "device_flow", "msal_app"]:
    if key not in st.session_state:
        st.session_state[key] = None
if "search_results" not in st.session_state:
    st.session_state.search_results = []
if "analysis_result" not in st.session_state:
    st.session_state.analysis_result = ""

# --- MSALアプリの初期化 ---
def get_msal_app():
    if st.session_state.msal_app is None:
        authority = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
        st.session_state.msal_app = msal.PublicClientApplication(
            MS_CLIENT_ID, authority=authority
        )
    return st.session_state.msal_app

# --- AIモデル自動選択 ---
def get_working_model():
    try:
        available_models = [
            m.name for m in genai.list_models()
            if 'generateContent' in m.supported_generation_methods
        ]
        target = next((m for m in available_models if "1.5-flash" in m), None)
        if not target:
            target = next((m for m in available_models if "flash" in m), None)
        if not target:
            target = available_models[0] if available_models else "gemini-1.5-flash"
        return genai.GenerativeModel(target)
    except Exception:
        return genai.GenerativeModel("gemini-1.5-flash")

# --- ファイルテキスト抽出 ---
def extract_text_from_url(download_url, file_name, token):
    headers = {'Authorization': f'Bearer {token}'}
    res = requests.get(download_url, headers=headers)
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
                    row_data = " ".join([str(c) for c in row if c is not None])
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

# --- Graph API ヘルパー ---
def graph_get(url, token):
    headers = {'Authorization': f'Bearer {token}'}
    res = requests.get(url, headers=headers)
    if res.status_code == 200:
        return res.json()
    return None

# --- Teamsチャンネル＆チャット一覧取得 ---
def get_teams_and_channels(token):
    items = []
    # 参加しているTeams一覧
    teams_data = graph_get("https://graph.microsoft.com/v1.0/me/joinedTeams", token)
    if teams_data:
        for team in teams_data.get('value', []):
            team_id = team['id']
            team_name = team['displayName']
            ch_data = graph_get(
                f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels", token
            )
            if ch_data:
                for ch in ch_data.get('value', []):
                    items.append({
                        'label': f"📢 {team_name} / {ch['displayName']}",
                        'type': 'channel',
                        'team_id': team_id,
                        'channel_id': ch['id'],
                    })
    # チャット一覧
    chat_data = graph_get("https://graph.microsoft.com/v1.0/me/chats?$expand=members", token)
    if chat_data:
        for chat in chat_data.get('value', []):
            chat_id = chat['id']
            members = chat.get('members', [])
            names = [m.get('displayName', '') for m in members if m.get('displayName')]
            label = "、".join(names[:3]) if names else chat_id
            items.append({
                'label': f"💬 {label}",
                'type': 'chat',
                'chat_id': chat_id,
            })
    return items

# --- チャンネルのファイル検索 ---
def search_channel_files(team_id, channel_id, keyword, token):
    results = []
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/filesFolder"
    data = graph_get(url, token)
    if not data:
        return results
    drive_id = data.get('parentReference', {}).get('driveId')
    item_id = data.get('id')
    if not drive_id or not item_id:
        return results
    search_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/search(q='{keyword}')"
    search_data = graph_get(search_url, token)
    if search_data:
        for item in search_data.get('value', []):
            if 'file' in item:
                results.append({
                    'name': item['name'],
                    'webUrl': item.get('webUrl', ''),
                    'downloadUrl': item.get('@microsoft.graph.downloadUrl', ''),
                })
    return results

# --- チャットの共有ファイル検索 ---
def search_chat_files(chat_id, keyword, token):
    results = []
    url = f"https://graph.microsoft.com/v1.0/me/chats/{chat_id}/messages?$top=50"
    data = graph_get(url, token)
    if not data:
        return results
    for msg in data.get('value', []):
        for att in msg.get('attachments', []):
            name = att.get('name', '')
            if keyword.lower() in name.lower():
                results.append({
                    'name': name,
                    'webUrl': att.get('contentUrl', ''),
                    'downloadUrl': '',
                })
    return results

# ======================
# UI
# ======================
st.title("📂 Teams ファイル検索アプリ")

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
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        st.session_state.ms_token = result["access_token"]
        st.rerun()

if not st.session_state.ms_token:
    if st.session_state.device_flow is None:
        if st.button("Microsoft 365 でログイン"):
            flow = app.initiate_device_flow(scopes=SCOPES)
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

# --- メイン機能 ---
if st.session_state.ms_token:
    token = st.session_state.ms_token

    st.header("② 検索先を選ぶ")
    with st.spinner("Teams・チャット一覧を取得中..."):
        channels = get_teams_and_channels(token)

    if not channels:
        st.warning("Teams・チャットが見つかりませんでした。")
    else:
        options = {ch['label']: ch for ch in channels}
        selected_label = st.selectbox("検索したいチャンネル・チャットを選んでください", list(options.keys()))
        selected = options[selected_label]

        st.header("③ キーワードで検索")
        keyword = st.text_input("検索キーワードを入力してください")

        if st.button("🔍 検索"):
            if not keyword:
                st.warning("キーワードを入力してください。")
            else:
                with st.spinner("検索中..."):
                    if selected['type'] == 'channel':
                        results = search_channel_files(
                            selected['team_id'], selected['channel_id'], keyword, token
                        )
                    else:
                        results = search_chat_files(selected['chat_id'], keyword, token)
                    st.session_state.search_results = results
                    st.session_state.analysis_result = ""

        if st.session_state.search_results:
            st.header("④ 解析するファイルを選択")
            st.write(f"{len(st.session_state.search_results)} 件見つかりました")

            selected_files = []
            for i, item in enumerate(st.session_state.search_results):
                checked = st.checkbox(item['name'], key=f"file_{i}")
                if item['webUrl']:
                    st.markdown(f"　🔗 [{item['webUrl']}]({item['webUrl']})")
                if checked:
                    selected_files.append(item)

            if st.button("🤖 AIで解析"):
                if not selected_files:
                    st.warning("ファイルを1つ以上選択してください。")
                else:
                    analysis_context = ""
                    with st.spinner("ファイルを読み取り中..."):
                        for item in selected_files:
                            if item['downloadUrl']:
                                text = extract_text_from_url(
                                    item['downloadUrl'], item['name'], token
                                )
                                if text:
                                    analysis_context += f"\n--- {item['name']} ---\n{text}\n"

                    if not analysis_context:
                        st.warning("テキストを抽出できませんでした。")
                    else:
                        with st.spinner("AIが分析中..."):
                            model = get_working_model()
                            prompt = (
                                f"あなたは優秀なアシスタントです。以下のファイルから"
                                f"「{keyword}」に関連する具体的な数値、日付、指示、変更点を抜き出して要約してください。\n\n"
                                f"【解析対象データ】\n{analysis_context}"
                            )
                            try:
                                ai_res = model.generate_content(prompt)
                                st.session_state.analysis_result = ai_res.text.strip()
                            except Exception as e:
                                st.error(f"AI分析エラー: {e}")

        if st.session_state.analysis_result:
            st.header("⑤ AIの分析結果")
            st.markdown(st.session_state.analysis_result)
