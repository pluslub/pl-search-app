import io
import msal
import requests
import streamlit as st
import google.generativeai as genai
from docx import Document
from datetime import datetime
import openpyxl
import re

# --- 設定情報 ---
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
defaults = {
    "ms_token": None,
    "device_flow": None,
    "msal_app": None,
    "channels_list": None,
    "message_results": [],
    "file_results": [],
    "analysis_result": "",
    "last_keyword": "",
}
for key, val in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = val

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

# --- Graph API ヘルパー ---
def graph_get(url, token):
    headers = {'Authorization': f'Bearer {token}'}
    res = requests.get(url, headers=headers)
    if res.status_code == 200:
        return res.json()
    return None

# --- HTMLタグを除去 ---
def strip_html(text):
    if not text:
        return ""
    return re.sub(r'<[^>]+>', '', text).strip()

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

# --- Teamsチャンネル＆チャット一覧取得 ---
def get_teams_and_channels(token):
    items = []
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
    chat_data = graph_get("https://graph.microsoft.com/v1.0/me/chats?$expand=members", token)
    if chat_data:
        for chat in chat_data.get('value', []):
            chat_id = chat['id']
            members = chat.get('members', [])
            names = [m.get('displayName', '') for m in members if m.get('displayName')]
            label = "、".join(names[:3]) if names else chat_id[:20]
            items.append({
                'label': f"💬 {label}",
                'type': 'chat',
                'chat_id': chat_id,
            })
    return items

# --- チャンネルのメッセージ検索 ---
def search_channel_messages(team_id, channel_id, keyword, token):
    results = []
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages?$top=50"
    data = graph_get(url, token)
    if not data:
        return results
    for msg in data.get('value', []):
        body = strip_html(msg.get('body', {}).get('content', ''))
        if keyword.lower() in body.lower():
            sender = msg.get('from', {})
            user = sender.get('user', {}) if sender else {}
            name = user.get('displayName', '不明') if user else '不明'
            created = msg.get('createdDateTime', '')
            try:
                dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
                date_str = dt.strftime('%Y/%m/%d %H:%M')
            except Exception:
                date_str = created
            results.append({
                'type': 'message',
                'sender': name,
                'date': date_str,
                'body': body[:500],
            })
    # 返信も検索
    for msg in data.get('value', []):
        msg_id = msg.get('id')
        reply_url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{msg_id}/replies?$top=50"
        reply_data = graph_get(reply_url, token)
        if reply_data:
            for reply in reply_data.get('value', []):
                body = strip_html(reply.get('body', {}).get('content', ''))
                if keyword.lower() in body.lower():
                    sender = reply.get('from', {})
                    user = sender.get('user', {}) if sender else {}
                    name = user.get('displayName', '不明') if user else '不明'
                    created = reply.get('createdDateTime', '')
                    try:
                        dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
                        date_str = dt.strftime('%Y/%m/%d %H:%M')
                    except Exception:
                        date_str = created
                    results.append({
                        'type': 'message',
                        'sender': name,
                        'date': date_str,
                        'body': body[:500],
                    })
    return results

# --- チャットのメッセージ検索 ---
def search_chat_messages(chat_id, keyword, token):
    results = []
    url = f"https://graph.microsoft.com/v1.0/me/chats/{chat_id}/messages?$top=50"
    data = graph_get(url, token)
    if not data:
        return results
    for msg in data.get('value', []):
        body = strip_html(msg.get('body', {}).get('content', ''))
        if keyword.lower() in body.lower():
            sender = msg.get('from', {})
            user = sender.get('user', {}) if sender else {}
            name = user.get('displayName', '不明') if user else '不明'
            created = msg.get('createdDateTime', '')
            try:
                dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
                date_str = dt.strftime('%Y/%m/%d %H:%M')
            except Exception:
                date_str = created
            results.append({
                'type': 'message',
                'sender': name,
                'date': date_str,
                'body': body[:500],
            })
        # 添付ファイルも収集
        for att in msg.get('attachments', []):
            att_name = att.get('name', '')
            if att_name:
                results.append({
                    'type': 'file',
                    'name': att_name,
                    'webUrl': att.get('contentUrl', ''),
                    'downloadUrl': '',
                })
    return results

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
                    'type': 'file',
                    'name': item['name'],
                    'webUrl': item.get('webUrl', ''),
                    'downloadUrl': item.get('@microsoft.graph.downloadUrl', ''),
                })
    return results

# ======================
# UI
# ======================
st.title("📂 Teams 横断検索アプリ")
st.caption("チャンネル・チャットのメッセージとファイルをキーワードで検索し、AIで要約します")

# --- Microsoft認証 ---
st.header("① Microsoft 365 にログイン")
app = get_msal_app()

if st.session_state.ms_token:
    st.success("✅ ログイン済みです")
    if st.button("ログアウト"):
        for key in defaults:
            st.session_state[key] = defaults[key]
        st.rerun()

if not st.session_state.ms_token:
    if st.session_state.device_flow is None:
        if st.button("Microsoft 365 でログイン"):
            flow = app.initiate_device_flow(scopes=SCOPES)
            st.session_state.device_flow = flow
            st.rerun()
    else:
        flow = st.session_state.device_flow
        st.info(
            f"以下のURLにアクセスしてコードを入力してください：\n\n"
            f"**{flow['verification_uri']}**\n\n"
            f"コード：**{flow['user_code']}**"
        )
        if st.button("ログイン完了（認証後にクリック）"):
            with st.spinner("認証確認中..."):
                result = app.acquire_token_by_device_flow(flow)
                if result and "access_token" in result:
                    st.session_state.ms_token = result["access_token"]
                    st.session_state.device_flow = None
                    st.rerun()
                else:
                    st.error("❌ 認証に失敗しました。もう一度お試しください。")
                    st.session_state.device_flow = None

# --- メイン機能 ---
if st.session_state.ms_token:
    token = st.session_state.ms_token

    # チャンネル一覧をキャッシュ
    if st.session_state.channels_list is None:
        with st.spinner("Teams・チャット一覧を取得中..."):
            st.session_state.channels_list = get_teams_and_channels(token)

    channels = st.session_state.channels_list

    if not channels:
        st.warning("Teams・チャットが見つかりませんでした。")
    else:
        st.header("② 検索先を選ぶ")
        labels = [ch['label'] for ch in channels]
        selected_index = st.selectbox(
            "検索したいチャンネル・チャットを選んでください",
            range(len(labels)),
            format_func=lambda i: labels[i],
        )
        selected = channels[selected_index]

        st.header("③ キーワードで検索")
        keyword = st.text_input("検索キーワードを入力してください")

        if st.button("🔍 メッセージ＆ファイルを検索"):
            if not keyword:
                st.warning("キーワードを入力してください。")
            else:
                st.session_state.last_keyword = keyword
                st.session_state.analysis_result = ""

                with st.spinner("メッセージを検索中..."):
                    if selected['type'] == 'channel':
                        msgs = search_channel_messages(
                            selected['team_id'], selected['channel_id'], keyword, token
                        )
                    else:
                        msgs = search_chat_messages(selected['chat_id'], keyword, token)

                msg_results = [m for m in msgs if m['type'] == 'message']
                file_from_chat = [m for m in msgs if m['type'] == 'file']

                with st.spinner("ファイルを検索中..."):
                    if selected['type'] == 'channel':
                        file_results = search_channel_files(
                            selected['team_id'], selected['channel_id'], keyword, token
                        )
                    else:
                        file_results = file_from_chat

                st.session_state.message_results = msg_results
                st.session_state.file_results = file_results

        # --- メッセージ検索結果 ---
        if st.session_state.message_results:
            st.header("④ メッセージ検索結果")
            st.write(f"💬 {len(st.session_state.message_results)} 件のメッセージが見つかりました")
            for i, msg in enumerate(st.session_state.message_results):
                with st.expander(f"📝 {msg['sender']}（{msg['date']}）"):
                    st.write(msg['body'])

        # --- ファイル検索結果 ---
        if st.session_state.file_results:
            st.header("⑤ ファイル検索結果")
            st.write(f"📁 {len(st.session_state.file_results)} 件のファイルが見つかりました")
            selected_files = []
            for i, item in enumerate(st.session_state.file_results):
                checked = st.checkbox(item['name'], key=f"file_{i}")
                if item.get('webUrl'):
                    st.markdown(f"　🔗 [{item['webUrl']}]({item['webUrl']})")
                if checked:
                    selected_files.append(item)

        # --- 検索結果なし ---
        if (st.session_state.last_keyword
                and not st.session_state.message_results
                and not st.session_state.file_results):
            st.info("検索結果が見つかりませんでした。")

        # --- AI解析ボタン ---
        if st.session_state.message_results or st.session_state.file_results:
            st.header("⑥ AIで要約・分析")
            if st.button("🤖 AIで分析する"):
                analysis_context = ""
                keyword = st.session_state.last_keyword

                # メッセージをコンテキストに追加
                for msg in st.session_state.message_results:
                    analysis_context += (
                        f"\n--- メッセージ({msg['sender']} / {msg['date']}) ---\n"
                        f"{msg['body']}\n"
                    )

                # 選択されたファイルをコンテキストに追加
                for item in st.session_state.file_results:
                    if item.get('downloadUrl'):
                        text = extract_text_from_url(
                            item['downloadUrl'], item['name'], token
                        )
                        if text:
                            analysis_context += f"\n--- ファイル: {item['name']} ---\n{text}\n"

                if not analysis_context:
                    st.warning("分析するデータがありません。")
                else:
                    with st.spinner("AIが分析中..."):
                        model = get_working_model()
                        prompt = (
                            f"あなたは優秀なアシスタントです。以下のメッセージとファイルの内容から、"
                            f"「{keyword}」に関連する情報を整理してください。\n"
                            f"・誰がいつ発言したか\n"
                            f"・具体的な数値、日付、指示、変更点\n"
                            f"・全体の要約\n\n"
                            f"【解析対象データ】\n{analysis_context}"
                        )
                        try:
                            ai_res = model.generate_content(prompt)
                            st.session_state.analysis_result = ai_res.text.strip()
                        except Exception as e:
                            st.error(f"AI分析エラー: {e}")

        # --- 分析結果表示 ---
        if st.session_state.analysis_result:
            st.header("📊 AIの分析結果")
            st.markdown(st.session_state.analysis_result)
