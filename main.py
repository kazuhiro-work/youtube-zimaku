import os
import re
import io
from google.colab import auth
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# ==========================================
# ■ 設定エリア (ユーザー指定情報)
# ==========================================
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1FUz828c016rg9xpqBThtaxv13M6XaafyjpZsG54bTws/edit?gid=0#gid=0"
FOLDER_ID = "1z7Tk3L5xCw6a71fpB0oVff_43GVEldRf"
CLIENT_SECRET_FILE = 'client_secret.json'

# ==========================================
# ■ 認証スコープ設定
# ==========================================
SCOPES_MAIN = [
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/spreadsheets'
]
SCOPES_BRAND = [
    'https://www.googleapis.com/auth/youtube.force-ssl'
]

# ==========================================
# ■ ヘルパー関数
# ==========================================
def authenticate_user(token_file, scopes, account_name_for_prompt):
    creds = None
    if os.path.exists(token_file):
        try:
            creds = Credentials.from_authorized_user_file(token_file, scopes)
        except:
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except:
                creds = None
        
        if not creds:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, 
                scopes,
                redirect_uri='http://localhost'
            )
            auth_url, _ = flow.authorization_url(prompt='consent')
            
            print(f"\n======== 【{account_name_for_prompt}】 の認証をお願いします ========")
            print("1. 以下のURLをクリックしてログインしてください。")
            print(f"2. ログイン画面では必ず **{account_name_for_prompt}** を選択してください。")
            print(auth_url)
            print("======================================================================")
            
            response_url = input(f"認証後のlocalhostのURLを貼り付けてEnter ({account_name_for_prompt}): ")
            
            try:
                code = re.search(r"code=([^&]+)", response_url).group(1)
                flow.fetch_token(code=code)
                creds = flow.credentials
                
                with open(token_file, 'w') as token:
                    token.write(creds.to_json())
                print(f"✅ {account_name_for_prompt} の認証が完了しました！\n")
            except Exception as e:
                print(f"❌ 認証エラー: URLが正しくない可能性があります。\n{e}")
                return None

    return creds

def clean_filename(text):
    return re.sub(r'[\\/:*?"<>|]', '-', text)

# --- [新規追加] 字幕クリーニング関数 (WebVTT → 簡易SBV風) ---
def clean_vtt_to_sbv_style(vtt_text):
    # 1. ヘッダー削除
    text = vtt_text.replace("WEBVTT\n", "").replace("WEBVTT", "")
    
    # 2. タイムスタンプの整形とメタデータの削除
    # 00:00:04.810 --> 00:00:08.850 align:start ... -> 0:00:04.810,0:00:08.850
    def format_timestamp(match):
        start = match.group(1).lstrip('0')
        if not start or start.startswith(':'): start = '0' + start
        end = match.group(2).lstrip('0')
        if not end or end.startswith(':'): end = '0' + end
        return f"\n{start},{end}"

    text = re.sub(r"(\d{2}:\d{2}:\d{2}\.\d{3}) --> (\d{2}:\d{2}:\d{2}\.\d{3}).*", format_timestamp, text)
    
    # 3. 単語レベルのタイムスタンプタグ <00:00:00.000> や <c> タグを削除
    text = re.sub(r"<[^>]+>", "", text)
    
    # 4. 行の整理と重複排除 (ASR特有の繰り返しを抑制)
    lines = text.split('\n')
    cleaned_lines = []
    for line in lines:
        s = line.strip()
        if not s: continue
        if not cleaned_lines or s != cleaned_lines[-1]:
            cleaned_lines.append(s)
            
    return "\n".join(cleaned_lines)

# ==========================================
# ■ メイン処理
# ==========================================
def main():
    print("🚀 Youtube字幕アーカイブ・システムを起動します...")
    
    # 1. 認証処理
    creds_main = authenticate_user('token_main.json', SCOPES_MAIN, "メインのGoogleアカウント")
    if not creds_main: return
    drive_service = build('drive', 'v3', credentials=creds_main)
    sheets_service = build('sheets', 'v4', credentials=creds_main)

    creds_brand = authenticate_user('token_brand.json', SCOPES_BRAND, "ブランドアカウント(YouTubeチャンネル)")
    if not creds_brand: return
    youtube_service = build('youtube', 'v3', credentials=creds_brand)

    # 2. シート読み込み
    try:
        sheet_id = re.search(r"/d/([^/]+)", SPREADSHEET_URL).group(1)
        result = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range="A2:G2000").execute()
        rows = result.get('values', [])
    except Exception as e:
        print(f"❌ スプレッドシート読み込みエラー: {e}")
        return

    # ★テスト用に3件で停止するように設定しています
    check_count = 0
    CHECK_LIMIT = 3 

    print(f"\n📋 データ処理を開始します (上限: {CHECK_LIMIT}件)")

    for i, row in enumerate(rows):
        if len(row) >= 6 and row[5]: 
            continue 
        
        if check_count >= CHECK_LIMIT:
            print("\n🛑 指定件数に達しました。")
            break

        date = row[0] if len(row) > 0 else "不明な日付"
        title = row[1] if len(row) > 1 else "タイトルなし"
        url = row[2] if len(row) > 2 else ""
        
        if not url: continue

        try:
            video_id = url.split('v=')[-1].split('&')[0]
        except:
            print(f"⚠ URL形式エラーのためスキップ: {url}")
            continue

        check_count += 1
        print(f"[{check_count}] 処理中: {title}")
        
        try:
            # --- [ブランド権限] 字幕を探す ---
            captions = youtube_service.captions().list(part='id,snippet', videoId=video_id).execute()
            
            if not captions.get('items'):
                print("   -> ⚠ 字幕データなし (スプレッドシートに記録します)")
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=sheet_id,
                    range=f"F{i+2}",
                    valueInputOption="RAW",
                    body={"values": [["字幕データなし"]]}
                ).execute()
                continue
            
            items = captions['items']
            target = next((c for c in items if c['snippet']['language'] == 'ja' and c['snippet']['trackKind'] != 'ASR'), None)
            if not target:
                target = next((c for c in items if c['snippet']['language'] == 'ja' and c['snippet']['trackKind'] == 'ASR'), None)
            if not target:
                target = items[0]
            
            print(f"   -> 字幕取得開始: {target['snippet']['trackKind']}")

            # --- [ブランド権限] ダウンロード ---
            req = youtube_service.captions().download(id=target['id'], tfmt='vtt')
            subtitle_content = req.execute().decode('utf-8')
            
            # --- [新規加筆] クリーニング実行 ---
            subtitle_content = clean_vtt_to_sbv_style(subtitle_content)
            
            # --- [メイン権限] ドライブ保存 ---
            raw_filename = f"{date}_{title}"
            safe_filename = clean_filename(raw_filename) + ".txt"
            
            file_metadata = {
                'name': safe_filename, 
                'parents': [FOLDER_ID]
            }
            
            media = MediaIoBaseUpload(io.BytesIO(subtitle_content.encode('utf-8')), mimetype='text/plain')
            file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            file_id = file.get('id')
            
            # --- [メイン権限] スプレッドシート更新 ---
            sheets_service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=f"F{i+2}",
                valueInputOption="RAW",
                body={"values": [[file_id]]}
            ).execute()
            
            print(f"   ✅ 保存成功 (ファイル名: {safe_filename})")
            
        except Exception as e:
            if "quotaExceeded" in str(e):
                print("   ❌ 本日のAPI制限に達しました。処理を中断します。")
                break
            print(f"   ❌ エラー: {e}")

if __name__ == "__main__":
    main()
