import sys
import os
import re
import gspread
from google.oauth2.service_account import Credentials

def parse_toml(content):
    # Simple TOML parser to avoid library dependency issues
    config = {}
    current_section = None
    for line in content.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("[") and line.endswith("]"):
            current_section = line[1:-1].strip()
            config[current_section] = {}
            continue
        if "=" in line and current_section:
            key, val = line.split("=", 1)
            key = key.strip()
            val = val.strip()
            # Strip quotes
            if val.startswith('"') and val.endswith('"'):
                val = val[1:-1]
            elif val.startswith("'") and val.endswith("'"):
                val = val[1:-1]
            config[current_section][key] = val
    return config

def main():
    try:
        print("1. Loading .streamlit/secrets.toml...")
        secrets_path = os.path.join(".streamlit", "secrets.toml")
        if not os.path.exists(secrets_path):
            print("[FAIL] .streamlit/secrets.toml 파일이 없습니다.")
            sys.exit(1)
            
        with open(secrets_path, "r", encoding="utf-8") as f:
            content = f.read()
            
        config = parse_toml(content)
        
        if "gcp_service_account" not in config or "google_sheets" not in config:
            print("[FAIL] [gcp_service_account] 또는 [google_sheets] 섹션이 올바르지 않습니다.")
            sys.exit(1)
            
        creds_dict = config["gcp_service_account"]
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            
        print("2. Authenticating with Google APIs...")
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(credentials)
        print("[SUCCESS] API 인증 성공!")
        
        print("3. Connecting to Spreadsheet...")
        spreadsheet_id = config["google_sheets"]["spreadsheet_id"]
        worksheet_name = config["google_sheets"]["worksheet_name"]
        print(f"   Spreadsheet ID: {spreadsheet_id}")
        print(f"   Worksheet Name: {worksheet_name}")
        
        spreadsheet = client.open_by_key(spreadsheet_id)
        try:
            worksheet = spreadsheet.worksheet(worksheet_name)
        except gspread.exceptions.WorksheetNotFound:
            print(f"\n[FAIL] '{worksheet_name}' 이름의 시트 탭을 찾을 수 없습니다.")
            print("현재 스프레드시트에 존재하는 시트 탭 목록:")
            for ws in spreadsheet.worksheets():
                print(f" - {ws.title}")
            sys.exit(1)
        print("[SUCCESS] 스프레드시트 연결 성공!")
        
        print("4. Fetching worksheet data...")
        all_values = worksheet.get_all_values()
        print(f"[SUCCESS] 데이터 가져오기 성공! 총 {len(all_values)}개의 행이 존재합니다.")
        if len(all_values) > 0:
            print("   헤더(1행):", all_values[0])
        else:
            print("   [WARNING] 경고: 시트가 비어있습니다.")
            
    except Exception as e:
        print(f"\n[FAIL] 오류 발생: {e}")
        print("스프레드시트가 아래 이메일에 공유되어 있는지 꼭 확인하세요:")
        try:
            print(f"Service Account Email: {config['gcp_service_account']['client_email']}")
        except:
            pass
        sys.exit(1)

if __name__ == "__main__":
    main()
