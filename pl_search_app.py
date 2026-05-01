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
    "evidence_links": [],
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
                        'team_name': team_name,
                        'channel_id': ch['id'],
                        'channel_name': ch['displayName'],
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

def get_channel_messages(team_id, channel_id, team_name, channel_name, token):
    all_data = []
    links = []
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages?$top=50"
    data = graph_get(url, token)
    if not data:
        return all_data, links

    for msg in data.get('value', []):
        raw_body = msg.get('body', {}).get('content', '')
        body = strip_html(raw_body)
        sender = msg.get('from', {})
        user = sender.get('user', {}) if sender else {}
        name = user.get('displayName', '不明') if user else '不明'
        created = msg.get('createdDateTime', '')
        msg_id = msg.get('id', '')
        try:
            dt = datetime.fromisoformat(created.replace('Z', '+00:00'))
            date_str = dt.strftime('%Y/%m/%d %H:%M')
        except Exception:
            date_str = created

        # TeamsメッセージへのディープリンクURL
        teams_link = f"https://teams.microsoft.com/l/message/{channel_id}/{msg_id}?groupId={team_id}&tenantId={MS_TENANT_ID}"

        attachments = msg.get('attachments', [])
        att_names = [a.get('name', '') for a in attachments if a.get('name')]

        entry = f"[メッセージID:{msg_id}] {name}（{date_str}）: {body}"
        if att_names:
            entry += f" [添付: {', '.join(att_names)}]"

        if body or att_names:
            all_data.append(entry[:500])
            links.append({
                'id': msg_id,
                'type': 'message',
                'label': f"{name}（{date_str}）",
                'url': teams_link,
                'detail': body[:100] or ', '.join(att_names),
            })

        # 返信
        reply_url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{msg_id}/replies?$top=20"
        reply_data = graph_get(reply_url, token)
        if reply_data:
            for reply in reply_data.get('value', []):
                rbody = strip_html(reply.get('body', {}).get('content', ''))
                rsender = reply.get('from', {})
                ruser = rsender.get('user', {}) if rsender else {}
                rname = ruser.get('displayName', '不明') if ruser else '不明'
                rcreated = reply.get('createdDateTime', '')
                reply_id = reply.get('id', '')
                try:
                    rdt = datetime.fromisoformat(rcreated.replace('Z', '+00:00'))
                    rdate_str = rdt.strftime('%Y/%m/%d %H:%M')
                except Exception:
                    rdate_str = rcreated
                ratts = reply.get('attachments', [])
                ratt_names = [a.get('name', '') for a in ratts if a.get('name')]
                rentry = f"[メッセージID:{reply_id}] {rname}（{rdate_str}）: {rbody}"
                if ratt_names:
                    rentry += f" [添付: {', '.join(ratt_names)}]"
                if rbody or ratt_names:
                    all_data.append(rentry[:500])
                    reply_link = f"https://teams.microsoft.com/l/message/{channel_id}/{reply_id}?groupId={team_id}&tenantId={MS_TENANT_ID}"
                    links.append({
                        'id': reply_id,
                        'type': 'message',
                        'label': f"{rname}（{rdate_str}）",
                        'url': reply_link,
                        'detail': rbody[:100] or ', '.join(ratt_names),
                    })

    return all_data, links

def get_channel_files(team_id, channel_id, token):
    all_data = []
    links = []
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/filesFolder"
    data = graph_get(url, token)
    if not data:
        return all_data, links
    drive_id = data.get('parentReference', {}).get('driveId')
    item_id = data.get('id')
    if not drive_id or not item_id:
        return all_data, links
    children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children?$top=50"
    children_data = graph_get(children_url, token)
    if not children_data:
        return all_data, links
    for item in children_data.get('value', []):
        if 'file' not in item:
            continue
        name = item['name']
        web_url = item.get('webUrl', '')
        file_item_id = item['id']
        file_drive_id = item.get('parentReference', {}).get('driveId', drive_id)
        links.append({
            'id': file_item_id,
            'type': 'file',
            'label': name,
            'url': web_url,
            'detail': name,
        })
        if not name.endswith(('.xlsx', '.xlsm', '.docx', '.txt')):
            all_data.append(f"[ファイルID:{file_item_id}] ファイル: {name}（{web_url}）: ※未対応形式")
            continue
        content = download_file_content(file_drive_id, file_item_id, token)
        if content:
            text = extract_text_from_bytes(content, name)
            if text:
                all_data.append(f"[ファイルID:{file_item_id}] ファイル: {name}（{web_url}）:\n{text[:2000]}")
    return all_data, links

def get_channel_onenote(team_id, token):
    all_data = []
    links = []
    pages_url = f"https://graph.microsoft.com/v1.0/groups/{team_id}/onenote/pages?$top=50&$select=id,title,createdDateTime,links"
    pages_data = graph_get(pages_url, token)
    if not pages_data:
        return all_data, links
    for page in pages_data.get('value', []):
        page_id = page.get('id', '')
        title = page.get('title', '無題')
        created = page.get('createdDateTime', '')
        # OneNoteページのリンク
        page_links = page.get('links', {})
        one_note_url = page_links.get('oneNoteWebUrl', {}).get('href', '')
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
                all_data.append(f"[OneNoteID:{page_id}] OneNote: {title}（{date_str}）:\n{body[:3000]}")
                links.append({
                    'id': page_id,
                    'type': 'onenote',
                    'label': f"OneNote: {title}（{date_str}）",
                    'url': one_note_url,
                    'detail': title,
                })
    return all_data, links

# ======================
# UI
# ======================
st.title("🔍 Teams AI検索アシスタント")
st.caption("メッセージ・ファイル・OneNoteを横断検索し、AIがエビデンス付きで回答します")

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
            placeholder="例：角谷さんが興味ありそうな求人は？"
        )

        if st.button("🚀 AIに聞く"):
            if not selected_indices:
                st.warning("検索先を選んでください。")
            elif not question:
                st.warning("質問を入力してください。")
            else:
                st.session_state.ai_answer = ""
                st.session_state.evidence_links = []
                all_context = []
                all_links = []

                progress = st.progress(0)
                total = len(selected_indices)

                for idx, sel_i in enumerate(selected_indices):
                    sel = channels[sel_i]
                    st.write(f"📡 {sel['label']} のデータを取得中...")

                    if sel['type'] == 'channel':
                        msgs, msg_links = get_channel_messages(
                            sel['team_id'], sel['channel_id'],
                            sel.get('team_name', ''), sel.get('channel_name', ''), token
                        )
                        all_context.extend(msgs)
                        all_links.extend(msg_links)

                        files, file_links = get_channel_files(sel['team_id'], sel['channel_id'], token)
                        all_context.extend(files)
                        all_links.extend(file_links)

                        notes, note_links = get_channel_onenote(sel['team_id'], token)
                        all_context.extend(notes)
                        all_links.extend(note_links)

                    progress.progress((idx + 1) / total)

                st.session_state.evidence_links = all_links

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
                            f"・回答には必ずエビデンス（いつ・誰が・どのOneNote/ファイル/メッセージ）を付けてください\n"
                            f"・エビデンスにはデータ内のID（OneNoteID・ファイルID・メッセージID）を必ず含めてください\n"
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

        # --- AI回答表示 ---
        if st.session_state.ai_answer:
            st.header("📊 AIの回答")
            st.markdown(st.session_state.ai_answer)

            # --- 記録資料リンク一覧 ---
            if st.session_state.evidence_links:
                st.header("🔗 記録資料一覧（元の記録へのリンク）")
                
                # 回答に登場したIDに対応するリンクを表示
                answer_text = st.session_state.ai_answer
                shown_links = []
                for link in st.session_state.evidence_links:
                    if link['id'] in answer_text and link['url']:
                        shown_links.append(link)
                
                # 登場したリンクがあれば表示、なければ全部表示
                display_links = shown_links if shown_links else st.session_state.evidence_links[:20]
                
                for link in display_links:
                    icon = "📝" if link['type'] == 'message' else "📄" if link['type'] == 'file' else "📓"
                    if link['url']:
                        st.markdown(f"{icon} [{link['label']}]({link['url']})")
                    else:
                        st.markdown(f"{icon} {link['label']}")
