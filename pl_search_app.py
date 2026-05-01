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
    "ai_answer": "",
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

# --- ファイルのコンテンツをダウンロード ---
def download_file_content(drive_id, item_id, token):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    headers = {'Authorization': f'Bearer {token}'}
    res = requests.get(url, headers=headers, allow_redirects=True)
    if res.status_code == 200:
        return res.content
    return None

# --- ファイルテキスト抽出 ---
def extract_text_from_bytes(file_bytes_raw, file_name):
    file_bytes = io.BytesIO(file_bytes_raw)
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
            text = file_bytes_raw.decode('utf-8', errors='ignore')
    except Exception as e:
        pass
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

# --- チャンネルのメッセージ全取得 ---
def get_channel_messages(team_id, channel_id, token, progress_cb=None):
    all_data = []
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages?$top=50"
    data = graph_get(url, token)
    if not data:
        return all_data
    for msg in data.get('value', []):
        body = strip_html(msg.get('body', {}).get('content', ''))
        if not body:
            continue
        sender = msg.get('from', {})
        user = sender.get('user', {}) if sender else {}
        name = user.get('displayName', '不明') if user else '不明'
        created = msg.get('createdDateTime', '')
        try:
            dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
            date_str = dt.strftime('%Y/%m/%d %H:%M')
        except Exception:
            date_str = created
        all_data.append(f"[メッセージ] {name}（{date_str}）: {body[:300]}")

        # 返信も取得
        msg_id = msg.get('id')
        reply_url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{msg_id}/replies?$top=20"
        reply_data = graph_get(reply_url, token)
        if reply_data:
            for reply in reply_data.get('value', []):
                rbody = strip_html(reply.get('body', {}).get('content', ''))
                if not rbody:
                    continue
                rsender = reply.get('from', {})
                ruser = rsender.get('user', {}) if rsender else {}
                rname = ruser.get('displayName', '不明') if ruser else '不明'
                rcreated = reply.get('createdDateTime', '')
                try:
                    rdt = datetime.fromisoformat(rcreated.replace('Z', '+00:00'))
                    rdate_str = rdt.strftime('%Y/%m/%d %H:%M')
                except Exception:
                    rdate_str = rcreated
                all_data.append(f"[返信] {rname}（{rdate_str}）: {rbody[:300]}")
    return all_data

# --- チャットのメッセージ全取得 ---
def get_chat_messages(chat_id, token):
    all_data = []
    url = f"https://graph.microsoft.com/v1.0/me/chats/{chat_id}/messages?$top=50"
    data = graph_get(url, token)
    if not data:
        return all_data
    for msg in data.get('value', []):
        body = strip_html(msg.get('body', {}).get('content', ''))
        if not body:
            continue
        sender = msg.get('from', {})
        user = sender.get('user', {}) if sender else {}
        name = user.get('displayName', '不明') if user else '不明'
        created = msg.get('createdDateTime', '')
        try:
            dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
            date_str = dt.strftime('%Y/%m/%d %H:%M')
        except Exception:
            date_str = created
        all_data.append(f"[メッセージ] {name}（{date_str}）: {body[:300]}")
    return all_data

# --- チャンネルのファイル全取得 ---
def get_channel_files(team_id, channel_id, token):
    all_data = []
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/filesFolder"
    data = graph_get(url, token)
    if not data:
        return all_data
    drive_id = data.get('parentReference', {}).get('driveId')
    item_id = data.get('id')
    if not drive_id or not item_id:
        return all_data

    # フォルダ内のファイル一覧
    children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children?$top=30"
    children_data = graph_get(children_url, token)
    if not children_data:
        return all_data

    for item in children_data.get('value', []):
        if 'file' not in item:
            continue
        name = item['name']
        web_url = item.get('webUrl', '')
        file_item_id = item['id']
        file_drive_id = item.get('parentReference', {}).get('driveId', drive_id)

        # 対応ファイル形式のみ
        if not name.endswith(('.xlsx', '.xlsm', '.docx', '.txt')):
            all_data.append(f"[ファイル] {name}（{web_url}）: ※未対応形式のためスキップ")
            continue

        content = download_file_content(file_drive_id, file_item_id, token)
        if content:
            text = extract_text_from_bytes(content, name)
            if text:
                all_data.append(f"[ファイル] {name}（{web_url}）:\n{text[:2000]}")
    return all_data

# ======================
# UI
# ======================
st.title("🔍 Teams AI検索アシスタント")
st.caption("選んだチャンネルのメッセージとファイルを元に、AIが質問に答えます")

# --- Microsoft認証 ---
app = get_msal_app()

if st.session_state.ms_token:
    st.success("✅ Microsoft 365 ログイン済み")
    col1, col2 = st.columns([6, 1])
    with col2:
        if st.button("ログアウト"):
            for key in defaults:
                st.session_state[key] = defaults[key]
            st.rerun()

if not st.session_state.ms_token:
    st.header("① Microsoft 365 にログイン")
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
        # --- 検索先選択（複数選択） ---
        labels = [ch['label'] for ch in channels]
        selected_indices = st.multiselect(
            "📂 検索先を選んでください（複数選択可）",
            range(len(labels)),
            format_func=lambda i: labels[i],
        )

        # --- 質問入力 ---
        question = st.text_input(
            "💬 質問を入力してください",
            placeholder="例：Aさんが体調不良だったのはいつ？"
        )

        if st.button("🚀 AIに聞く"):
            if not selected_indices:
                st.warning("検索先を1つ以上選んでください。")
            elif not question:
                st.warning("質問を入力してください。")
            else:
                st.session_state.ai_answer = ""
                all_context = []

                # 進捗表示
                progress = st.progress(0)
                total = len(selected_indices)

                for idx, sel_i in enumerate(selected_indices):
                    sel = channels[sel_i]
                    st.write(f"📡 {sel['label']} のデータを取得中...")

                    if sel['type'] == 'channel':
                        msgs = get_channel_messages(
                            sel['team_id'], sel['channel_id'], token
                        )
                        files = get_channel_files(
                            sel['team_id'], sel['channel_id'], token
                        )
                        all_context.extend(msgs)
                        all_context.extend(files)
                    else:
                        msgs = get_chat_messages(sel['chat_id'], token)
                        all_context.extend(msgs)

                    progress.progress((idx + 1) / total)

                if not all_context:
                    st.warning("データが取得できませんでした。")
                else:
                    # データ量を制限（Geminiの入力上限対策）
                    context_text = "\n".join(all_context)
                    if len(context_text) > 50000:
                        context_text = context_text[:50000] + "\n...(以下省略)"

                    with st.spinner("🤖 AIが分析中..."):
                        model = get_working_model()
                        prompt = (
                            f"あなたは社内アシスタントです。以下のTeamsメッセージとファイルの内容を元に、"
                            f"ユーザーの質問に正確に答えてください。\n\n"
                            f"【ルール】\n"
                            f"・回答には必ず根拠（エビデンス）を付けてください\n"
                            f"・エビデンスは「いつ・誰が・どのファイル/メッセージで」の形式で示してください\n"
                            f"・該当する情報が複数あれば、すべて時系列で列挙してください\n"
                            f"・見つからない場合は「該当する情報は見つかりませんでした」と回答してください\n\n"
                            f"【ユーザーの質問】\n{question}\n\n"
                            f"【検索対象データ】\n{context_text}"
                        )
                        try:
                            ai_res = model.generate_content(prompt)
                            st.session_state.ai_answer = ai_res.text.strip()
                        except Exception as e:
                            st.error(f"AI分析エラー: {e}")

        # --- AI回答表示 ---
        if st.session_state.ai_answer:
            st.header("📊 AIの回答")
            st.markdown(st.session_state.ai_answer)
