import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

st.title("🔗 Google Sheets 연동 테스트")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

try:
    st.write("### 1. Streamlit secrets 로딩 시도 중...")
    creds_dict = dict(st.secrets["gcp_service_account"])
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    
    st.success("✅ Secrets 로딩 성공!")
    
    st.write("### 2. Google API 인증 시도 중...")
    credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(credentials)
    st.success("✅ Google API 인증 성공!")
    
    st.write("### 3. 스프레드시트 열기 시도 중...")
    spreadsheet_id = st.secrets["google_sheets"]["spreadsheet_id"]
    worksheet_name = st.secrets["google_sheets"]["worksheet_name"]
    st.write(f"- Spreadsheet ID: `{spreadsheet_id}`")
    st.write(f"- Worksheet Name: `{worksheet_name}`")
    
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet(worksheet_name)
    st.success("✅ 스프레드시트 및 워크시트 접근 성공!")
    
    st.write("### 4. 기존 데이터 조회 테스트...")
    all_values = worksheet.get_all_values()
    st.success(f"✅ 데이터 조회 성공! (총 {len(all_values)}개의 행 존재)")
    
    if len(all_values) > 0:
        st.write("- 헤더(1행):", all_values[0])
    else:
        st.warning("⚠️ 시트가 완전히 비어있습니다. 헤더가 설정되어 있어야 합니다.")
        
except Exception as e:
    st.error(f"❌ 오류 발생: {e}")
    st.info("비밀키가 올바르게 입력되었는지, 스프레드시트가 해당 서비스 계정 이메일에 공유(편집자 권한)되어 있는지 다시 확인해 주세요.")
