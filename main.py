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
# メインアカウント用 (ドライブ保存・スプレッドシート書き込み)
SCOPES_MAIN = [
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/spreadsheets'
]
# ブランドアカウント用 (YouTube字幕取得)
SCOPES_BRAND = [
    'https://www.googleapis.com/auth/youtube.force-ssl'
]

# ==========================================
# ■ ヘルパー関数
# ==========================================
# 認証処理 (トークンファイルを使い分ける)
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

# ファイル名クリーニング (Win/Mac/Linuxで禁止されている文字を置換)
def clean_filename(text):
    # スラッシュ、コロン、アスタリスク、クエスチョン、引用符、不等号、パイプをハイフンに
    return re.sub(r'[\\/:*?"<>|]', '-', text)

# ==========================================
# ■ メイン処理
# ==========================================
def main():
    print("🚀 Youtube字幕アーカイブ・システムを起動します...")
    
    # ---------------------------------------------------------
    # 1. メインアカウント認証 (ドライブ・スプレッドシート用)
    # ---------------------------------------------------------
    creds_main = authenticate_user('token_main.json', SCOPES_MAIN, "メインのGoogleアカウント")
    if not creds_main: return
    drive_service = build('drive', 'v3', credentials=creds_main)
    sheets_service = build('sheets', 'v4', credentials=creds_main)

    # ---------------------------------------------------------
    # 2. ブランドアカウント認証 (YouTube用)
    # ---------------------------------------------------------
    creds_brand = authenticate_user('token_brand.json', SCOPES_BRAND, "ブランドアカウント(YouTubeチャンネル)")
    if not creds_brand: return
    youtube_service = build('youtube', 'v3', credentials=creds_brand)

    # ---------------------------------------------------------
    # 3. データ処理開始
    # ---------------------------------------------------------
    try:
        sheet_id = re.search(r"/d/([^/]+)", SPREADSHEET_URL).group(1)
        # データ範囲を取得 (ヘッダーを除く2行目から)
        result = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range="A2:G2000").execute()
        rows = result.get('values', [])
    except Exception as e:
        print(f"❌ スプレッドシート読み込みエラー: {e}")
        return

    check_count = 0
    CHECK_LIMIT = 3 # ★1回あたりの処理上限 (必要に応じて変更)

    print(f"\n📋 データ処理を開始します (今回の上限: {CHECK_LIMIT}件)")

    for i, row in enumerate(rows):
        # F列(インデックス5)に既にIDがある場合はスキップ (完了済み)
        if len(row) >= 6 and row[5]: 
            continue 
        
        # 安全停止チェック
        if check_count >= CHECK_LIMIT:
            print("\n🛑 指定件数に達しました。")
            break

        # データ取得 (A列:日付, B列:タイトル, C列:URL)
        date = row[0] if len(row) > 0 else "不明な日付"
        title = row[1] if len(row) > 1 else "タイトルなし"
        url = row[2] if len(row) > 2 else ""
        
        if not url: continue # URLがない行は無視

        # 動画ID抽出
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
                print("   -> ⚠ 字幕データなし")
                continue
            
            items = captions['items']
            # 日本語の手動字幕 -> 日本語のASR(自動) -> なければ先頭 の順で選択
            target = next((c for c in items if c['snippet']['language'] == 'ja' and c['snippet']['trackKind'] != 'ASR'), None)
            if not target:
                target = next((c for c in items if c['snippet']['language'] == 'ja' and c['snippet']['trackKind'] == 'ASR'), None)
            if not target:
                target = items[0]
            
            print(f"   -> 字幕取得開始: {target['snippet']['trackKind']}")

            # --- [ブランド権限] ダウンロード ---
            req = youtube_service.captions().download(id=target['id'], tfmt='vtt')
            subtitle_content = req.execute().decode('utf-8')
            
            # --- [メイン権限] ファイル名作成とドライブ保存 ---
            # 仕様: 投稿日_動画タイトル.txt (特殊文字はハイフン化)
            raw_filename = f"{date}_{title}"
            safe_filename = clean_filename(raw_filename) + ".txt"
            
            file_metadata = {
                'name': safe_filename, 
                'parents': [FOLDER_ID]
            }
            
            media = MediaIoBaseUpload(io.BytesIO(subtitle_content.encode('utf-8')), mimetype='text/plain')
            
            file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            file_id = file.get('id')
            
            # --- [メイン権限] スプレッドシート更新 (F列にID記載) ---
            # i + 2 で実際の行番号を指定 (A2スタートのため)
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
