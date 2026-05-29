import json
import os
import sys

def main():
    if len(sys.argv) < 3:
        print("사용법: python convert_credentials.py <json_파일_경로> <스프레드시트_ID>")
        print("예시: python convert_credentials.py my-key.json 1ABC_spreadsheet_id_XYZ")
        sys.exit(1)

    json_path = sys.argv[1]
    spreadsheet_id = sys.argv[2]

    if not os.path.exists(json_path):
        print(f"[FAIL] JSON 파일을 찾을 수 없습니다: {json_path}")
        sys.exit(1)

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            creds = json.load(f)
    except Exception as e:
        print(f"[FAIL] JSON 파싱 실패: {e}")
        sys.exit(1)

    # 필수 필드 확인
    required_keys = ["project_id", "private_key_id", "private_key", "client_email", "client_id"]
    for key in required_keys:
        if key not in creds:
            print(f"[FAIL] 올바른 서비스 계정 키 파일이 아닙니다. '{key}' 필드가 없습니다.")
            sys.exit(1)

    # private_key 가공 (줄바꿈 문자를 \\n 으로 치환하여 한 줄로 만듬)
    private_key = creds["private_key"].replace("\n", "\\n")

    # .streamlit 폴더 및 secrets.toml 작성
    os.makedirs(".streamlit", exist_ok=True)
    secrets_path = os.path.join(".streamlit", "secrets.toml")

    toml_content = f"""# ============================================================
# Streamlit secrets.toml (자동 생성됨)
# ============================================================

[gcp_service_account]
type = "service_account"
project_id = "{creds.get('project_id', '')}"
private_key_id = "{creds.get('private_key_id', '')}"
private_key = "{private_key}"
client_email = "{creds.get('client_email', '')}"
client_id = "{creds.get('client_id', '')}"
auth_uri = "{creds.get('auth_uri', 'https://accounts.google.com/o/oauth2/auth')}"
token_uri = "{creds.get('token_uri', 'https://oauth2.googleapis.com/token')}"
auth_provider_x509_cert_url = "{creds.get('auth_provider_x509_cert_url', 'https://www.googleapis.com/oauth2/v1/certs')}"
client_x509_cert_url = "{creds.get('client_x509_cert_url', '')}"

[google_sheets]
spreadsheet_id = "{spreadsheet_id}"
worksheet_name = "VOC관리대장"
"""

    try:
        with open(secrets_path, "w", encoding="utf-8") as f:
            f.write(toml_content)
        print(f"[SUCCESS] 연동용 secrets 설정 파일이 성공적으로 생성되었습니다: {secrets_path}")
        print(f"[INFO] 서비스 계정 이메일: {creds['client_email']}")
        print("[INFO] 위 이메일 주소를 구글 스프레드시트에 '편집자' 권한으로 꼭 공유해 주세요!")
    except Exception as e:
        print(f"[FAIL] secrets.toml 쓰기 실패: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
