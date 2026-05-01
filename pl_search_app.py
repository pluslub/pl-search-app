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
    "Notes.Read.All",
]

defaults = {
    "ms_token": None,
    "device_flow": None,
    "msal_app": None,
    "channels_list": None,
    "ai_answer": "",
    "debug_info": "",
    "raw_messages": [],
}
for key, val in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = val

def get_msal_app():
    if st.session_state.msal_app is None:
        authority = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
        st.session_state.msal_app = msal.PublicClientApplication(
            MS_CLIENT_ID, authority=authority
        )
    return st.session_state.msal_app

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

def graph_get(url, token):
    headers = {'Authorization': f'Bearer {token}'}
    res = requests.get(url, headers=headers)
    if res.status_code == 200:
        return res.json()
    return None

def strip_html(text):
    if not text:
        return ""
    return re.sub(r'<[^>]+>', '', text).strip()

def download_file_content(drive_id, item_id, token):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    headers = {'Authorization': f'Bearer {token}'}
    res = requests.get(url, headers=headers, allow_redirects=True)
    if res.status_code == 200:
        return res.content
    return None

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
        text = f"(解析エラー: {e})"
    return text[:4000]

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

def get_channel_messages(team_id, channel_id, token):
    all_data = []
    raw = []
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages?$top=50"
    data = graph_get(url, token)
    if not data:
        return all_data, raw

    for msg in data.get('value', []):
        # 生のbody content
        raw_body = msg.get('body', {}).get('content', '')
        body = strip_html(raw_body)

        sender = msg.get('from', {})
        user = sender.get('user', {}) if sender else {}
        name = user.get('displayName', '不明') if user else '不明'
        created = msg.get('createdDateTime', '')
        try:
            dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
            date_str = dt.strftime('%Y/%m/%d %H:%M')
        except Exception:
            date_str = created

        # 添付ファイル名もテキストとして追加
        attachments = msg.get('attachments', [])
        att_names = [a.get('name', '') for a in attachments if a.get('name')]

        entry = f"[メッセージ] {name}（{date_str}）: {body}"
        if att_names:
            entry += f" [添付: {', '.join(att_names)}]"

        if body or att_names:
            all_data.append(entry[:500])
            raw.append({'sender': name, 'date': date_str, 'body': body[:200], 'attachments': att_names})

        # 返信
        msg_id = msg.get('id')
        reply_url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{msg_id}/replies?$top=20"
        reply_data = graph_get(reply_url, token)
        if reply_data:
            for reply in reply_data.get('value', []):
                rbody = strip_html(reply.get('body', {}).get('content', ''))
                rsender = reply.get('from', {})
                ruser = rsender.get('user', {}) if rsender else {}
                rname = ruser.get('displayName', '不明') if ruser else '不明'
                rcreated = reply.get('createdDateTime', '')
                try:
                    rdt = datetime.fromisoformat(rcreated.replace('Z', '+00:00'))
                    rdate_str = rdt.strftime('%Y/%m/%d %H:%M')
                except Exception:
                    rdate_str = rcreated
                ratts = reply.get('attachments', [])
                ratt_names = [a.get('name', '') for a in ratts if a.get('name')]
                rentry = f"[返信] {rname}（{rdate_str}）: {rbody}"
                if ratt_names:
                    rentry += f" [添付: {', '.join(ratt_names)}]"
                if rbody or ratt_names:
                    all_data.append(rentry[:500])
                    raw.append({'sender': rname, 'date': rdate_str, 'body': rbody[:200], 'attachments': ratt_names})

    return all_data, raw

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
    children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children?$top=50"
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
        if not name.endswith(('.xlsx', '.xlsm', '.docx', '.txt')):
            all_data.append(f"[ファイル: {name}]（{web_url}）: ※未対応形式")
            continue
        content = download_file_content(file_drive_id, file_item_id, token)
        if content:
            text = extract_text_from_bytes(content, name)
            if text:
                all_data.append(f"[ファイル: {name}]（{web_url}）:\n{text[:2000]}")
    return all_data

def get_channel_onenote(team_id, token):
    all_data = []
    pages_url = f"https://graph.microsoft.com/v1.0/groups/{team_id}/onenote/pages?$top=50&$select=id,title,createdDateTime"
    pages_data = graph_get(pages_url, token)
    if not pages_data:
        return all_data
    for page in pages_data.get('value', []):
        page_id = page.get('id', '')
        title = page.get('title', '無題')
        created = page.get('createdDateTime', '')
        try:
            dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
            date_str = dt.strftime('%Y/%m/%d')
        except Exception:
            date_str = created
        content_url = f"https://graph.microsoft.com/v1.0/groups/{team_id}/onenote/pages/{page_id}/content"
        headers = {'Authorization': f'Bearer {token}'}
        res = requests.get(content_url, headers=headers)
        if res.status_code == 200:
            body = strip_html(res.text)
            if body.strip():
                all_data.append(f"[OneNote: {title}]（{date_str}）:\n{body[:3000]}")
    return all_data

# ======================
# UI
# ======================
st.title("🔍 Teams AI検索アシスタント")

app = get_msal_app()

if st.session_state.ms_token:
    st.success("✅ ログイン済み")
    col1, col2 = st.columns([6, 1])
    with col2:
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
                    st.error("❌ 認証に失敗しました。")
                    st.session_state.device_flow = None

if st.session_state.ms_token:
    token = st.session_state.ms_token

    if st.session_state.channels_list is None:
        with st.spinner("Teams・チャット一覧を取得中..."):
            st.session_state.channels_list = get_teams_and_channels(token)

    channels = st.session_state.channels_list

    if channels:
        labels = [ch['label'] for ch in channels]
        selected_indices = st.multiselect(
            "📂 検索先を選んでください（複数選択可）",
            range(len(labels)),
            format_func=lambda i: labels[i],
        )

        question = st.text_input(
            "💬 質問を入力してください",
            placeholder="例：小島さんについて教えて"
        )

        if st.button("🚀 AIに聞く"):
            if not selected_indices:
                st.warning("検索先を選んでください。")
            elif not question:
                st.warning("質問を入力してください。")
            else:
                st.session_state.ai_answer = ""
                st.session_state.raw_messages = []
                all_context = []
                debug_lines = []

                progress = st.progress(0)
                total = len(selected_indices)

                for idx, sel_i in enumerate(selected_indices):
                    sel = channels[sel_i]
                    st.write(f"📡 {sel['label']} のデータを取得中...")

                    if sel['type'] == 'channel':
                        msgs, raw = get_channel_messages(sel['team_id'], sel['channel_id'], token)
                        st.session_state.raw_messages.extend(raw)
                        debug_lines.append(f"{sel['label']}: メッセージ {len(msgs)} 件")
                        all_context.extend(msgs)

                        files = get_channel_files(sel['team_id'], sel['channel_id'], token)
                        debug_lines.append(f"{sel['label']}: ファイル {len(files)} 件")
                        all_context.extend(files)

                        notes = get_channel_onenote(sel['team_id'], token)
                        debug_lines.append(f"{sel['label']}: OneNote {len(notes)} 件")
                        all_context.extend(notes)

                    progress.progress((idx + 1) / total)

                total_chars = len("\n".join(all_context))
                debug_lines.append(f"合計データ量: {total_chars} 文字")
                st.session_state.debug_info = "\n".join(debug_lines)

                if not all_context:
                    st.warning("データが取得できませんでした。")
                else:
                    context_text = "\n".join(all_context)
                    if len(context_text) > 50000:
                        context_text = context_text[:50000]

                    with st.spinner("🤖 AIが分析中..."):
                        model = get_working_model()
                        prompt = (
                            f"あなたは社内アシスタントです。以下のデータを元に質問に答えてください。\n\n"
                            f"【ルール】\n"
                            f"・回答には必ずエビデンス（いつ・誰が・どのファイル）を付けてください\n"
                            f"・該当情報が複数あれば時系列で列挙してください\n"
                            f"・見つからない場合は「見つかりませんでした」と答えてください\n\n"
                            f"【質問】\n{question}\n\n"
                            f"【データ】\n{context_text}"
                        )
                        try:
                            ai_res = model.generate_content(prompt)
                            st.session_state.ai_answer = ai_res.text.strip()
                        except Exception as e:
                            st.error(f"AI分析エラー: {e}")

        if st.session_state.ai_answer:
            st.header("📊 AIの回答")
            st.markdown(st.session_state.ai_answer)

        if st.session_state.debug_info:
            with st.expander("🔧 取得データの内訳"):
                st.text(st.session_state.debug_info)

        # 取得したメッセージの中身を表示（デバッグ用）
        if st.session_state.raw_messages:
            with st.expander(f"📝 取得したメッセージ一覧（{len(st.session_state.raw_messages)}件）"):
                for msg in st.session_state.raw_messages[:20]:
                    st.write(f"**{msg['sender']}**（{msg['date']}）")
                    st.write(msg['body'] if msg['body'] else "（本文なし）")
                    if msg['attachments']:
                        st.write(f"📎 {', '.join(msg['attachments'])}")
                    st.divider()
