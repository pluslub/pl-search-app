import io
import msal
import requests
import streamlit as st
import google.generativeai as genai
from docx import Document
from datetime import datetime
import openpyxl
import re
from supabase import create_client

# --- 設定情報 ---
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
MS_CLIENT_ID = st.secrets["MS_CLIENT_ID"]
MS_TENANT_ID = st.secrets["MS_TENANT_ID"]
SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]

genai.configure(api_key=GEMINI_API_KEY)
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

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
        authority = "https://login.microsoftonline.com/" + MS_TENANT_ID
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
    headers = {'Authorization': 'Bearer ' + token}
    res = requests.get(url, headers=headers)
    if res.status_code == 200:
        return res.json()
    return None


def strip_html(text):
    if not text:
        return ""
    return re.sub(r'<[^>]+>', '', text).strip()


def download_file_content(drive_id, item_id, token):
    url = (
        "https://graph.microsoft.com/v1.0/drives/"
        + drive_id + "/items/" + item_id + "/content"
    )
    headers = {'Authorization': 'Bearer ' + token}
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
                text += "\n[シート: " + sheet.title + "]\n"
                for row in sheet.iter_rows(values_only=True):
                    row_data = " ".join([str(c) for c in row if c is not None])
                    if row_data.strip():
                        text += row_data + "\n"
        elif file_name.endswith('.docx'):
            doc = Document(file_bytes)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif file_name.endswith('.txt'):
            text = file_bytes_raw.decode('utf-8', errors='ignore')
        elif file_name.endswith('.pdf'):
            try:
                import pypdf
                pdf_reader = pypdf.PdfReader(io.BytesIO(file_bytes_raw))
                for page in pdf_reader.pages:
                    text += page.extract_text() or ""
            except Exception:
                text = "(PDF解析エラー)"
    except Exception as e:
        text = "(解析エラー: " + str(e) + ")"
    return text[:4000]


def get_embedding(text):
    url = (
        "https://generativelanguage.googleapis.com/v1/models/"
        "text-embedding-004:embedContent"
    )
    body = {
        "model": "models/text-embedding-004",
        "content": {"parts": [{"text": text[:2000]}]},
        "taskType": "RETRIEVAL_DOCUMENT"
    }
    res = requests.post(url, json=body, params={"key": GEMINI_API_KEY})
    res.raise_for_status()
    return res.json()["embedding"]["values"]


# --- Supabaseにドキュメントを保存 ---
def save_document(source_type, source_id, title, content, author, recorded_at, url, channel_name, team_name):
    if not content or not content.strip():
        return
    try:
        embedding = get_embedding(content)
        existing = supabase.table("documents").select("id").eq("source_id", source_id).execute()
        if existing.data:
            supabase.table("documents").update({
                "content": content,
                "embedding": embedding,
                "updated_at": datetime.now().isoformat(),
            }).eq("source_id", source_id).execute()
        else:
            supabase.table("documents").insert({
                "source_type": source_type,
                "source_id": source_id,
                "title": title,
                "content": content,
                "embedding": embedding,
                "author": author,
                "recorded_at": recorded_at,
                "url": url,
                "channel_name": channel_name,
                "team_name": team_name,
            }).execute()
    except Exception as e:
        st.warning("DB保存エラー: " + str(e))


# --- Supabaseからベクトル検索 ---
def search_documents(query_text, channel_names=None):
    try:
        embed_url = (
            "https://generativelanguage.googleapis.com/v1/models/"
            "text-embedding-004:embedContent"
        )
        embed_body = {
            "model": "models/text-embedding-004",
            "content": {"parts": [{"text": query_text}]},
            "taskType": "RETRIEVAL_QUERY"
        }
        embed_res = requests.post(embed_url, json=embed_body, params={"key": GEMINI_API_KEY})
        embed_res.raise_for_status()
        query_embedding = embed_res.json()["embedding"]["values"]
        result = supabase.rpc("match_documents", {
            "query_embedding": query_embedding,
            "match_threshold": 0.5,
            "match_count": 20,
            "filter_channels": channel_names or None,
        }).execute()
        return result.data or []
    except Exception as e:
        st.warning("DB検索エラー: " + str(e))
        return []


def get_teams_and_channels(token):
    items = []
    teams_data = graph_get("https://graph.microsoft.com/v1.0/me/joinedTeams", token)
    if teams_data:
        for team in teams_data.get('value', []):
            team_id = team['id']
            team_name = team['displayName']
            ch_url = "https://graph.microsoft.com/v1.0/teams/" + team_id + "/channels"
            ch_data = graph_get(ch_url, token)
            if ch_data:
                for ch in ch_data.get('value', []):
                    items.append({
                        'label': "📢 " + team_name + " / " + ch['displayName'],
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
                'label': "💬 " + label,
                'type': 'chat',
                'chat_id': chat_id,
            })
    return items


def index_channel(sel, token):
    team_id = sel['team_id']
    channel_id = sel['channel_id']
    team_name = sel.get('team_name', '')
    channel_name = sel.get('channel_name', '')
    count = 0

    msg_url = (
        "https://graph.microsoft.com/v1.0/teams/"
        + team_id + "/channels/" + channel_id + "/messages?$top=50"
    )
    data = graph_get(msg_url, token)
    if data:
        for msg in data.get('value', []):
            body = strip_html(msg.get('body', {}).get('content', ''))
            sender = msg.get('from', {})
            user = sender.get('user', {}) if sender else {}
            name = user.get('displayName', '不明') if user else '不明'
            created = msg.get('createdDateTime', '')
            msg_id = msg.get('id', '')
            teams_link = (
                "https://teams.microsoft.com/l/message/"
                + channel_id + "/" + msg_id
                + "?groupId=" + team_id
                + "&tenantId=" + MS_TENANT_ID
            )
            atts = msg.get('attachments', [])
            att_names = [a.get('name', '') for a in atts if a.get('name')]
            full_content = body
            if att_names:
                full_content += " [添付: " + ", ".join(att_names) + "]"
            if full_content.strip():
                save_document(
                    'message', msg_id, None, full_content,
                    name, created, teams_link, channel_name, team_name
                )
                count += 1

            reply_url = (
                "https://graph.microsoft.com/v1.0/teams/"
                + team_id + "/channels/" + channel_id
                + "/messages/" + msg_id + "/replies?$top=20"
            )
            reply_data = graph_get(reply_url, token)
            if reply_data:
                for reply in reply_data.get('value', []):
                    rbody = strip_html(reply.get('body', {}).get('content', ''))
                    rsender = reply.get('from', {})
                    ruser = rsender.get('user', {}) if rsender else {}
                    rname = ruser.get('displayName', '不明') if ruser else '不明'
                    rcreated = reply.get('createdDateTime', '')
                    reply_id = reply.get('id', '')
                    reply_link = (
                        "https://teams.microsoft.com/l/message/"
                        + channel_id + "/" + reply_id
                        + "?groupId=" + team_id
                        + "&tenantId=" + MS_TENANT_ID
                    )
                    if rbody.strip():
                        save_document(
                            'message', reply_id, None, rbody,
                            rname, rcreated, reply_link, channel_name, team_name
                        )
                        count += 1

    folder_url = (
        "https://graph.microsoft.com/v1.0/teams/"
        + team_id + "/channels/" + channel_id + "/filesFolder"
    )
    folder_data = graph_get(folder_url, token)
    if folder_data:
        drive_id = folder_data.get('parentReference', {}).get('driveId')
        item_id = folder_data.get('id')
        if drive_id and item_id:
            children_url = (
                "https://graph.microsoft.com/v1.0/drives/"
                + drive_id + "/items/" + item_id + "/children?$top=50"
            )
            children_data = graph_get(children_url, token)
            if children_data:
                for item in children_data.get('value', []):
                    if 'file' not in item:
                        continue
                    name = item['name']
                    web_url = item.get('webUrl', '')
                    file_item_id = item['id']
                    file_drive_id = item.get('parentReference', {}).get('driveId', drive_id)
                    supported = ('.xlsx', '.xlsm', '.docx', '.txt', '.pdf')
                    if not name.endswith(supported):
                        continue
                    content = download_file_content(file_drive_id, file_item_id, token)
                    if content:
                        text = extract_text_from_bytes(content, name)
                        if text:
                            save_document(
                                'file', file_item_id, name, text,
                                None, None, web_url, channel_name, team_name
                            )
                            count += 1

    pages_url = (
        "https://graph.microsoft.com/v1.0/groups/"
        + team_id + "/onenote/pages"
        + "?$top=50&$select=id,title,createdDateTime,links"
    )
    pages_data = graph_get(pages_url, token)
    if pages_data:
        for page in pages_data.get('value', []):
            page_id = page.get('id', '')
            title = page.get('title', '無題')
            created = page.get('createdDateTime', '')
            page_links = page.get('links', {})
            one_note_url = page_links.get('oneNoteWebUrl', {}).get('href', '')
            content_url = (
                "https://graph.microsoft.com/v1.0/groups/"
                + team_id + "/onenote/pages/" + page_id + "/content"
            )
            headers = {'Authorization': 'Bearer ' + token}
            res = requests.get(content_url, headers=headers)
            if res.status_code == 200:
                body = strip_html(res.text)
                if body.strip():
                    save_document(
                        'onenote', page_id, title, body,
                        None, created, one_note_url, channel_name, team_name
                    )
                    count += 1

    return count


# ======================
# UI
# ======================
st.title("🔍 Plusらぼ AI検索アシスタント")
st.caption("メッセージ・ファイル・OneNote・PDFを横断検索し、AIが関連情報をまとめて回答します")

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
            "以下のURLにアクセスしてコードを入力してください：\n\n"
            "**" + flow['verification_uri'] + "**\n\n"
            "コード：**" + flow['user_code'] + "**"
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

        tab1, tab2 = st.tabs(["🔍 検索", "🔄 インデックス更新"])

        with tab1:
            selected_indices = st.multiselect(
                "📂 検索先を選んでください（複数選択可）",
                range(len(labels)),
                format_func=lambda i: labels[i],
            )

            question = st.text_input(
                "💬 質問を入力してください",
                placeholder="例：Aさんの体調変化について"
            )

            if st.button("🚀 AIに聞く"):
                if not selected_indices:
                    st.warning("検索先を選んでください。")
                elif not question:
                    st.warning("質問を入力してください。")
                else:
                    st.session_state.ai_answer = ""
                    st.session_state.evidence_links = []

                    selected_channel_names = [
                        channels[i].get('channel_name') for i in selected_indices
                        if channels[i].get('channel_name')
                    ]

                    with st.spinner("DBから検索中..."):
                        all_docs = search_documents(question, selected_channel_names)

                    if not all_docs:
                        st.warning("DBにデータがありません。「インデックス更新」タブでデータを取り込んでください。")
                    else:
                        all_context = []
                        all_links = []
                        for doc in all_docs:
                            source_type = doc.get('source_type', '')
                            source_id = doc.get('source_id', '')
                            title = doc.get('title', '')
                            content = doc.get('content', '')
                            author = doc.get('author', '不明')
                            recorded_at = doc.get('recorded_at', '')
                            url = doc.get('url', '')

                            try:
                                dt = None
                                if recorded_at:
                                    dt = datetime.fromisoformat(recorded_at.replace('Z', '+00:00'))
                                date_str = dt.strftime('%Y/%m/%d %H:%M') if dt else ''
                            except Exception:
                                date_str = recorded_at

                            if source_type == 'message':
                                entry = "[メッセージID:" + source_id + "] " + author + "（" + date_str + "）: " + content[:500]
                                icon = "📝"
                                label = author + "（" + date_str + "）"
                            elif source_type == 'file':
                                entry = "[ファイルID:" + source_id + "] ファイル: " + title + ":\n" + content[:1000]
                                icon = "📄"
                                label = title or source_id
                            else:
                                entry = "[OneNoteID:" + source_id + "] OneNote: " + title + "（" + date_str + "）:\n" + content[:2000]
                                icon = "📓"
                                label = title + "（" + date_str + "）"

                            all_context.append(entry)
                            all_links.append({
                                'id': source_id,
                                'type': source_type,
                                'label': icon + " " + label,
                                'url': url,
                            })

                        st.session_state.evidence_links = all_links

                        context_text = "\n".join(all_context)
                        if len(context_text) > 50000:
                            context_text = context_text[:50000]

                        with st.spinner("🤖 AIが分析中..."):
                            model = get_working_model()
                            prompt = (
                                "あなたは福祉施設の支援記録を管理する社内アシスタントです。"
                                "以下のデータを元に質問に答えてください。\n\n"
                                "【重要なルール】\n"
                                "・質問のキーワードだけでなく、福祉現場で関連するあらゆる言葉・状況を幅広く拾ってください\n"
                                "・回答には必ず記録資料（いつ・誰が・どのOneNote/ファイル/メッセージ）のIDを含めてください\n"
                                "・該当情報が複数あれば時系列で列挙してください\n"
                                "・見つからない場合は「見つかりませんでした」と答えてください\n\n"
                                "【質問】\n" + question + "\n\n"
                                "【データ】\n" + context_text
                            )
                            try:
                                ai_res = model.generate_content(prompt)
                                st.session_state.ai_answer = ai_res.text.strip()
                            except Exception as e:
                                st.error("AI分析エラー: " + str(e))

            if st.session_state.ai_answer:
                st.header("📊 AIの回答")
                st.markdown(st.session_state.ai_answer)

                if st.session_state.evidence_links:
                    st.header("🔗 記録資料一覧")
                    answer_text = st.session_state.ai_answer
                    shown_links = [
                        link for link in st.session_state.evidence_links
                        if link['id'] in answer_text and link['url']
                    ]
                    display_links = shown_links if shown_links else [
                        lnk for lnk in st.session_state.evidence_links[:20] if lnk['url']
                    ]
                    for link in display_links:
                        st.markdown("[" + link['label'] + "](" + link['url'] + ")")

        with tab2:
            st.write("選択したチャンネルのデータをDBに取り込みます。初回や新しい記録が増えたときに実行してください。")

            index_indices = st.multiselect(
                "📂 インデックス化するチャンネルを選んでください",
                range(len(labels)),
                format_func=lambda i: labels[i],
                key="index_select"
            )

            if st.button("🔄 インデックス更新を実行"):
                if not index_indices:
                    st.warning("チャンネルを選んでください。")
                else:
                    total_count = 0
                    for sel_i in index_indices:
                        sel = channels[sel_i]
                        if sel['type'] != 'channel':
                            continue
                        with st.spinner(sel['label'] + " を取り込み中..."):
                            count = index_channel(sel, token)
                            total_count += count
                            st.write("✅ " + sel['label'] + ": " + str(count) + " 件保存しました")
                    st.success("🎉 合計 " + str(total_count) + " 件のデータをDBに保存しました！")
