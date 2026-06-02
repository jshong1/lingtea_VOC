"""
VOC 데일리 응대 입력 시스템
----------------------------
고객감동팀 VOC 상담/문의 건을 Google Sheets에 저장하는 Streamlit 웹 애플리케이션
"""

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import datetime
import re
import json
import anthropic
import streamlit.components.v1 as components

# ============================================================
# CS 분류표 (대 > 중 > 소 종속 드롭다운)
# 향후 Google Sheets의 마스터 시트에서 불러오는 구조로 확장 가능
# 확장 시: get_cs_categories_from_sheet() 함수를 추가하고
#           아래 CS_CATEGORIES 딕셔너리를 해당 함수 호출로 대체
# ============================================================
CS_CATEGORIES = {
    "주문": {
        "D2C주문": [],
        "B2B주문": [],
        "변경": ["고객정보변경", "옵션변경"],
        "교환": ["변심", "일정/절차", "소비기한도래"],
        "반품": ["변심", "일정/절차"],
        "취소/철회": ["취소", "철회"],
        "출고": ["배송일정", "분리배송"],
        "일반요청": ["결제/입금/확인", "기타"],
    },
    "문의": {
        "프로모션": ["공식몰/네이버", "온라인/오픈마켓/라이브", "오프라인", "기타", "전화주문"],
        "제품": ["효능/효과", "섭취문의", "성분/함량", "제품별비교", "판매처문의", "기타"],
        "협업": ["광고/제안", "협찬/기증/후원", "납품문의"],
        "공식몰": ["멤버쉽", "설문조사", "쿠폰/적립금", "서버/오류"],
        "협력사": ["약국", "홈쇼핑", "판매사이트", "택배사", "타부서/기타"],
    },
    "고객의견": {
        "상품": ["제안", "불만"],
        "서비스": ["상담응대불만", "미수긍"],
        "오류": [],
        "정책": ["리셀", "출고프로세스"],
    },
    "품질": {
        "관능": ["이취", "맛"],
        "이물": ["일반", "상해", "혐오"],
        "변성": ["색변화", "굳음", "융해"],
        "불량": ["실링불량", "용기불량", "수량/중량부족", "내용물없음", "보틀", "기타"],
        "부작용": ["소화기질환", "피부질환", "호흡기질환"],
        "고객": ["확인불가", "고객과실", "착오/기타"],
    },
    "배송": {
        "배송사": ["파손", "지연", "오배송", "분실", "기타"],
        "본사": ["출고누락", "오출고", "기타"],
        "물류사": ["출고누락", "오출고", "내품파손", "기타"],
        "고객": ["고객착오", "고객과실", "기타"],
    },
    "단순": {
        "등기수령": [],
        "기타": [],
        "문의내용없음": [],
        "재인입": ["일반문의", "품질", "항의/미수긍"],
    },
}

# ============================================================
# Google Sheets 연동 함수
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def get_gsheet_client():
    """Google Sheets API 클라이언트를 반환합니다."""
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        # private_key의 \\n을 실제 개행문자로 변환
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        client = gspread.authorize(credentials)
        return client
    except Exception as e:
        st.error(f"Google Sheets 인증 오류: {e}")
        return None


def get_worksheet():
    """설정된 스프레드시트의 워크시트를 반환합니다."""
    client = get_gsheet_client()
    if client is None:
        return None
    try:
        spreadsheet_id = st.secrets["google_sheets"]["spreadsheet_id"]
        worksheet_name = st.secrets["google_sheets"]["worksheet_name"]
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        return worksheet
    except Exception as e:
        st.error(f"워크시트 접근 오류: {e}")
        return None


# ============================================================
# 핵심 데이터 로딩 — API 호출은 단 1회 (10분 캐싱)
# load_all_records() 하나로 전체 시트를 가져온 뒤
# 판매처·유형·고객유형 목록은 그 결과에서 파생합니다.
# ============================================================

@st.cache_data(ttl=600)  # 10분 캐싱
def load_all_records():
    """구글 시트의 전체 VOC 데이터를 1회 API 호출로 가져와서 역순 리스트로 반환합니다."""
    worksheet = get_worksheet()
    if worksheet is None:
        return []
    try:
        all_values = worksheet.get_all_values()
        if len(all_values) <= 1:
            return []

        records_list = []
        total_rows = len(all_values)
        for idx in range(total_rows - 1, 0, -1):
            row = all_values[idx]
            if not row or not row[0]:
                continue

            row_dict = {
                "row_idx": idx + 1,  # gspread는 1-based index
                "연번":        row[0],
                "월번":        row[1]  if len(row) >  1 else "",
                "NO":          row[2]  if len(row) >  2 else "",
                "판매처":      row[3]  if len(row) >  3 else "",
                "담당자":      row[4]  if len(row) >  4 else "",
                "유형":        row[5]  if len(row) >  5 else "",
                "접수 일자":   row[6]  if len(row) >  6 else "",
                "고객유형":    row[7]  if len(row) >  7 else "",
                "고객명":      row[8]  if len(row) >  8 else "",
                "고객 전화번호": row[9] if len(row) >  9 else "",
                "대":          row[10] if len(row) > 10 else "",
                "중":          row[11] if len(row) > 11 else "",
                "소":          row[12] if len(row) > 12 else "",
                "문의내용":    row[13] if len(row) > 13 else "",
                "주문여부":    row[14] if len(row) > 14 else "",
                "금액":        row[15] if len(row) > 15 else "",
                "첫/재주문":   row[16] if len(row) > 16 else "",
                "연령대":      row[17] if len(row) > 17 else "",
                "성별":        row[18] if len(row) > 18 else "",
                "인입경로":    row[19] if len(row) > 19 else "",
                "중복 여부":   row[20] if len(row) > 20 else "",
                "재출고 여부": row[21] if len(row) > 21 else "",
                "성함":        row[22] if len(row) > 22 else "",
                "운송장":      row[23] if len(row) > 23 else "",
                "제품명":      row[24] if len(row) > 24 else "",
                "발생수량(EA)": row[25] if len(row) > 25 else "",
                "클레임 유형": row[26] if len(row) > 26 else "",
                "보상":        row[27] if len(row) > 27 else "",
            }
            records_list.append(row_dict)
        return records_list
    except Exception as e:
        st.error(f"데이터 로딩 실패: {e}")
        return []


# ── 아래 세 함수는 추가 API 호출 없이 캐시된 전체 데이터에서 파생합니다 ──

def load_unique_sellers():
    """판매처(D열) 고유 목록 — load_all_records() 캐시에서 파생."""
    default_sellers = ["공식몰", "네이버스마트스토어", "쿠팡", "지마켓", "카카오톡 선물하기"]
    records = load_all_records()
    if not records:
        return default_sellers
    exclude = {"판매처", "통화내역/온라인 접수내역", ""}
    sellers = {r["판매처"].strip() for r in records if r["판매처"].strip() not in exclude}
    for ds in default_sellers:
        sellers.add(ds)
    return sorted(sellers)


def load_unique_types():
    """유형(F열) 고유 목록 — load_all_records() 캐시에서 파생."""
    default_types = ["온라인", "유선"]
    records = load_all_records()
    if not records:
        return default_types
    exclude = {"유형", "통화내역/온라인 접수내역", ""}
    types = {r["유형"].strip() for r in records if r["유형"].strip() not in exclude}
    for dt in default_types:
        types.add(dt)
    return sorted(types)


def load_unique_customer_types():
    """고객유형(H열) 고유 목록 — load_all_records() 캐시에서 파생."""
    default_cust_types = ["일반고객", "강성고객", "단골고객"]
    records = load_all_records()
    if not records:
        return default_cust_types
    exclude = {"고객유형", "통화내역/온라인 접수내역", ""}
    cust_types = {r["고객유형"].strip() for r in records if r["고객유형"].strip() not in exclude}
    for dct in default_cust_types:
        cust_types.add(dct)
    return sorted(cust_types)


def get_next_numbers(worksheet, receipt_date: datetime.date):
    """
    연번, 월번, NO를 자동 생성합니다.

    - 연번: 기존 데이터의 연번 최대값 + 1
    - 월번: 접수 일자 기준 해당 월의 기존 건수 + 1
    - NO: 연번과 동일 (필요 시 별도 로직으로 분리 가능)
    """
    try:
        all_values = worksheet.get_all_values()

        # 헤더 행 제외
        data_rows = all_values[1:] if len(all_values) > 1 else []

        # 연번 계산 (컬럼 인덱스 0)
        max_serial = 0
        for row in data_rows:
            if row and row[0]:
                try:
                    val = int(str(row[0]).strip())
                    if val > max_serial:
                        max_serial = val
                except ValueError:
                    pass
        next_serial = max_serial + 1

        # 월번 계산 (컬럼 인덱스 6: 접수 일자)
        target_year = receipt_date.year
        target_month = receipt_date.month
        month_count = 0
        for row in data_rows:
            if len(row) > 6 and row[6]:
                date_str = str(row[6]).strip()
                try:
                    # YYYY-MM-DD 형식 파싱
                    row_date = datetime.datetime.strptime(date_str[:10], "%Y-%m-%d").date()
                    if row_date.year == target_year and row_date.month == target_month:
                        month_count += 1
                except ValueError:
                    pass
        next_monthly = month_count + 1

        # NO: 연번과 동일 (추후 별도 로직으로 변경 가능)
        next_no = next_serial

        return next_serial, next_monthly, next_no

    except Exception as e:
        st.error(f"번호 생성 오류: {e}")
        return 1, 1, 1


# ============================================================
# CS 분류 헬퍼 함수
# ============================================================

def get_large_categories():
    """대분류 목록을 반환합니다."""
    return list(CS_CATEGORIES.keys())


def get_middle_categories(large: str):
    """선택된 대분류에 해당하는 중분류 목록을 반환합니다."""
    if not large or large not in CS_CATEGORIES:
        return []
    return list(CS_CATEGORIES[large].keys())


def get_small_categories(large: str, middle: str):
    """선택된 대분류/중분류에 해당하는 소분류 목록을 반환합니다."""
    if not large or not middle:
        return []
    if large not in CS_CATEGORIES:
        return []
    mid_dict = CS_CATEGORIES[large]
    if middle not in mid_dict:
        return []
    return mid_dict[middle]  # 빈 리스트이면 소분류 없음


# ============================================================
# 유효성 검사
# ============================================================

def validate_form(data: dict, show_delivery: bool) -> list:
    """
    폼 유효성 검사 후 오류 메시지 리스트를 반환합니다.
    오류가 없으면 빈 리스트 반환.
    """
    errors = []

    # 기본 필수값
    required_basic = {
        "판매처": data.get("판매처"),
        "담당자": data.get("담당자"),
        "유형": data.get("유형"),
        "접수 일자": data.get("접수 일자"),
        "고객유형": data.get("고객유형"),
        "대": data.get("대"),
        "중": data.get("중"),
        "문의내용": data.get("문의내용"),
    }
    for field, value in required_basic.items():
        if not value or str(value).strip() == "":
            errors.append(f"'{field}' 항목은 필수 입력값입니다.")

    # 소분류: 소분류 옵션이 있는 경우에만 필수
    large = data.get("대", "")
    middle = data.get("중", "")
    small_options = get_small_categories(large, middle)
    if small_options and not data.get("소"):
        errors.append("'소' 항목은 필수 입력값입니다.")

    # 주문여부 O인 경우 추가 필수값
    if data.get("주문여부") == "O":
        if not data.get("금액") and data.get("금액") != 0:
            errors.append("주문여부가 O인 경우 '금액'은 필수 입력값입니다.")
        if not data.get("첫/재주문"):
            errors.append("주문여부가 O인 경우 '첫/재주문'은 필수 입력값입니다.")

    # 배송품질 정보 입력 시 추가 필수값
    if show_delivery:
        delivery_required = {
            "성함": data.get("성함"),
            "제품명": data.get("제품명"),
            "발생수량(EA)": data.get("발생수량(EA)"),
            "클레임 유형": data.get("클레임 유형"),
        }
        for field, value in delivery_required.items():
            if value is None or str(value).strip() == "":
                errors.append(f"배송품질 정보: '{field}' 항목은 필수 입력값입니다.")

    return errors


# ============================================================
# Google Sheets 저장
# ============================================================

def append_to_sheet(worksheet, row_data: list) -> bool:
    """Google Sheets에 행을 추가합니다."""
    try:
        worksheet.append_row(row_data, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"저장 오류: {e}")
        return False


def update_sheet(worksheet, row_idx: int, row_data: list) -> bool:
    """Google Sheets의 특정 행을 업데이트(수정)합니다."""
    try:
        # A{row_idx}부터 순서대로 업데이트 (예: A10:AB10)
        # gspread에서 list of list 형태로 보냄
        col_letter = chr(64 + len(row_data)) if len(row_data) <= 26 else "AB" # 최대 AB열까지 지원
        cell_range = f"A{row_idx}:{col_letter}{row_idx}"
        worksheet.update(cell_range, [row_data], value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"수정 오류: {e}")
        return False


# ============================================================
# 폼 초기화
# ============================================================

def reset_form(keep_shared_values: bool = False):
    """
    Session State를 초기화하여 폼을 리셋합니다.
    keep_shared_values가 True이면 공통 정보(판매처, 담당자, 유형, 접수일자, 분류 등)를 유지합니다.
    """
    # 초기화에서 제외할 항목 설정 (값 유지 모드일 때)
    shared_keys = [
        "판매처", "판매처_선택", "판매처_직접입력", 
        "유형", "유형_선택", "유형_직접입력",
        "고객유형", "고객유형_선택", "고객유형_직접입력",
        "담당자", "접수일자",
        "대분류", "중분류", "소분류"
    ]
    
    keys_to_reset = [
        "판매처", "판매처_선택", "판매처_직접입력", 
        "유형", "유형_선택", "유형_직접입력",
        "고객유형", "고객유형_선택", "고객유형_직접입력",
        "담당자", "접수일자",
        "고객명", "고객전화번호", "대분류", "중분류", "소분류",
        "문의내용", "주문여부", "금액", "첫재주문", "연령대",
        "성별", "인입경로", "중복여부", "show_delivery",
        "재출고여부", "성함", "운송장", "제품명", "발생수량",
        "클레임유형", "보상",
    ]
    
    for key in keys_to_reset:
        if keep_shared_values and key in shared_keys:
            continue
        if key in st.session_state:
            del st.session_state[key]
            
    # 수정 모드 상태도 해제
    if "edit_mode" in st.session_state:
        del st.session_state["edit_mode"]
    if "edit_row_idx" in st.session_state:
        del st.session_state["edit_row_idx"]
    if "edit_serial" in st.session_state:
        del st.session_state["edit_serial"]
        
    st.session_state["form_reset"] = True



# ============================================================
# AI 분석 리포트 생성 로직
# ============================================================

def generate_ai_report(records, selected_model):
    try:
        api_key = st.secrets["anthropic"]["api_key"]
    except KeyError:
        st.error("설정 오류: `.streamlit/secrets.toml` 파일에 `[anthropic]` 및 `api_key`가 설정되어 있지 않습니다.")
        return

    if not records:
        st.warning("분석할 VOC 데이터가 없습니다.")
        return

    # 전처리: 토큰 절약을 위해 핵심 필드만 추출
    processed_data = []
    for r in records:
        processed_data.append({
            "날짜": r.get("접수 일자", "")[:10],
            "채널": r.get("유형", ""),
            "분류": f"{r.get('대', '')}>{r.get('중', '')}>{r.get('소', '')}",
            "클레임": r.get("클레임 유형", ""),
            "내용": r.get("문의내용", "")
        })

    # JSON 덤프
    data_str = json.dumps(processed_data, ensure_ascii=False)
    
    # 데이터가 너무 크면 자르기 (대략 15만 자)
    if len(data_str) > 150000:
        st.warning("데이터가 너무 많아 최근 데이터 위주로 분석이 제한될 수 있습니다.")
        data_str = data_str[:150000]

    prompt = f"""
당신은 고객 경험(CX) 데이터 분석가입니다. 
다음은 특정 기간 동안 인입된 VOC(Voice of Customer) 데이터입니다.
총 {len(processed_data)}건의 데이터가 있습니다.

이 데이터를 바탕으로 종합적인 'VOC 분석 리포트'를 마크다운 형식으로 작성해주세요.
리포트는 다음 목차를 포함하여 체계적이고 구체적인 수치/내용을 바탕으로 작성해주시기 바랍니다:

1. **VOC 현황 요약**: 총 접수 건수 및 전반적 추이 (채널별 문의량, 주요 이슈 등)
2. **채널 및 유형별 분석**: 유입 채널별 비중 및 대/중/소 분류별 주요 문의 유형 비중
3. **고객 불만/제안 요약**: '내용' 필드에서 반복적으로 나타나는 주요 불만 사항(예: 파손, 환불, 품질 등) 및 고객 제안 사항 상세 요약
4. **품질/배송 민원 현황**: 관능/변성/이물 등 품질 관련 클레임과 파손/지연/오배송 등 배송 관련 클레임 현황 (구체적인 사례 포함)
5. **종합 인사이트 및 제안**: 데이터 기반 CX 향상 및 프로세스 개선을 위한 액션 아이템 제안

⚠️ **강력한 제한 사항 (글자수 초과/짤림 방지)**: 
- 데이터량이 많아 답변이 길어질 경우 텍스트가 중간에 짤리는 현상이 발생하고 있습니다.
- 따라서 **전체 리포트는 절대 2,500자를 넘지 않도록 극도로 요약**해 주세요.
- 각 목차(1~5번)별로 **최대 3~4개의 핵심 포인트만 불릿(Bullet) 형태**로 서술하고, 불필요한 부연 설명은 생략하여 문장이 끝까지 온전히 완성되도록 조절해 주세요.

VOC 데이터:
```json
{data_str}
```
"""

    with st.spinner(f"{selected_model} 모델이 데이터를 분석하고 리포트를 생성하는 중입니다..."):
        try:
            client = anthropic.Anthropic(api_key=api_key)
            response = client.messages.create(
                model=selected_model,
                max_tokens=4000,
                temperature=0.2,
                system="전문적인 CX 분석 리포트를 마크다운으로 작성하세요. 중요한 숫자나 주요 이슈는 굵은 글씨로 강조하세요.",
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            report = response.content[0].text
            st.markdown(report)
            
            # PDF 인쇄(장표 형태)를 위한 전역 CSS 주입
            print_css = """
            <style>
            @media print {
                /* 사이드바, 헤더 등 숨김 */
                [data-testid="stSidebar"], 
                header[data-testid="stHeader"], 
                .stApp > header, 
                footer, 
                .stDeployButton {
                    display: none !important;
                }
                
                /* 여백 최적화 */
                .stApp, .main .block-container {
                    max-width: 100% !important;
                    padding: 0 !important;
                    margin: 0 !important;
                }

                /* 장표 느낌의 폰트 및 페이지 나누기 */
                body {
                    font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif !important;
                    font-size: 14pt !important;
                    color: #000 !important;
                }
                
                /* 2번 항목(대제목 수준) 등에서 페이지 넘김 유도 */
                h1, h2 {
                    page-break-before: always;
                    break-before: page;
                    margin-top: 0 !important;
                }
                h1:first-of-type, h2:first-of-type {
                    page-break-before: auto;
                    break-before: auto;
                }

                p, li, table {
                    page-break-inside: avoid;
                    break-inside: avoid;
                }
            }
            </style>
            """
            st.markdown(print_css, unsafe_allow_html=True)
            
            # 인쇄를 트리거하는 자바스크립트 버튼 주입
            print_button_html = """
            <div style="text-align: right; margin-top: 20px;">
                <button onclick="window.parent.print()" 
                        style="padding: 12px 24px; background-color: #ff4b4b; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 16px; font-weight: bold; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                    📥 리포트 PDF로 다운로드 (인쇄)
                </button>
            </div>
            """
            components.html(print_button_html, height=80)
            
        except Exception as e:
            st.error(f"AI 리포트 생성 중 오류가 발생했습니다: {e}")

def render_ai_report_tab(records):
    st.title("📊 AI VOC 리포트 분석")
    st.write(f"현재 사이드바 필터 조건에 해당하는 **{len(records)}건**의 데이터를 바탕으로 AI 분석 리포트를 생성합니다.")
    st.info("💡 제공된 월간 리포트 양식(현황 요약, 채널 분석, 불만/제안 요약, 품질민원 등)을 기반으로 심층 분석을 수행합니다.")
    
    # 모델 선택 옵션 제공 (계정 권한/버전에 따른 404 방지)
    selected_model = st.selectbox(
        "사용할 AI 모델 선택 (오류 발생 시 다른 모델로 변경해보세요)",
        options=[
            "claude-sonnet-4-5",
            "claude-3-5-sonnet-20241022",
            "claude-3-5-sonnet-20240620",
            "claude-3-sonnet-20240229",
            "claude-3-haiku-20240307",
            "claude-3-opus-20240229"
        ],
        index=0
    )
    
    if st.button("🤖 AI 리포트 생성 시작", type="primary", use_container_width=True):
        generate_ai_report(records, selected_model)

# ============================================================
# Streamlit 앱 메인
# ============================================================

def main():
    st.set_page_config(
        page_title="VOC 데일리 응대 입력 시스템",
        page_icon="📋",
        layout="wide",
    )

    # --------------------------------------------------------
    # 사이드바: 날짜/판매처/담당자 필터링 및 불러오기/수정 기능
    # --------------------------------------------------------
    st.sidebar.title("📋 최근 등록 내역 조회")

    # 1. 날짜 필터 입력창 배치
    st.sidebar.markdown("### 📅 접수일자 필터")
    col_start, col_end = st.sidebar.columns(2)
    with col_start:
        start_date = st.date_input("시작일", value=datetime.date.today() - datetime.timedelta(days=7), key="filter_start")
    with col_end:
        end_date = st.date_input("종료일", value=datetime.date.today(), key="filter_end")

    # 2. 판매처(D열) / 담당자(E열) 필터 — 기존 데이터 기반 동적 목록
    st.sidebar.markdown("### 🏪 판매처 / 👤 담당자 필터")

    # 전체 데이터에서 판매처·담당자 고유값 추출 (필터 선택창 구성용)
    _all_for_filter = load_all_records()
    _all_sellers  = sorted({r["판매처"]  for r in _all_for_filter if r.get("판매처")})
    _all_managers = sorted({r["담당자"]  for r in _all_for_filter if r.get("담당자")})

    filter_sellers  = st.sidebar.multiselect(
        "판매처 (D열)",
        options=_all_sellers,
        default=[],
        placeholder="전체 (선택 안 하면 전체 조회)",
        key="filter_sellers",
    )
    filter_managers = st.sidebar.multiselect(
        "담당자 (E열)",
        options=_all_managers,
        default=[],
        placeholder="전체 (선택 안 하면 전체 조회)",
        key="filter_managers",
    )

    search_triggered = st.sidebar.button("조회 🔍", use_container_width=True)

    # 세션 상태로 조회 조건을 유지
    if "filter_start_applied" not in st.session_state:
        st.session_state["filter_start_applied"]   = start_date
        st.session_state["filter_end_applied"]     = end_date
        st.session_state["filter_sellers_applied"] = filter_sellers
        st.session_state["filter_managers_applied"]= filter_managers

    if search_triggered:
        st.session_state["filter_start_applied"]   = start_date
        st.session_state["filter_end_applied"]     = end_date
        st.session_state["filter_sellers_applied"] = filter_sellers
        st.session_state["filter_managers_applied"]= filter_managers

    # 데이터 로딩 및 필터링
    all_records = load_all_records()
    filtered_records = []

    # 날짜 + 판매처 + 담당자 필터 조건 복합 적용
    applied_sellers  = st.session_state.get("filter_sellers_applied",  [])
    applied_managers = st.session_state.get("filter_managers_applied", [])

    for rec in all_records:
        # G열(접수 일자) 날짜 필터
        rec_date_str = rec.get("접수 일자", "").strip()[:10]
        try:
            rec_date = datetime.datetime.strptime(rec_date_str, "%Y-%m-%d").date()
        except ValueError:
            continue
        if not (st.session_state["filter_start_applied"] <= rec_date <= st.session_state["filter_end_applied"]):
            continue

        # D열(판매처) 필터 — 선택된 값이 없으면 전체 허용
        if applied_sellers and rec.get("판매처", "") not in applied_sellers:
            continue

        # E열(담당자) 필터 — 선택된 값이 없으면 전체 허용
        if applied_managers and rec.get("담당자", "") not in applied_managers:
            continue

        filtered_records.append(rec)

    # G열(접수 일자) 기준 최신순 정렬
    def parse_rec_date(rec):
        d_str = rec.get("접수 일자", "").strip()[:10]
        try:
            return datetime.datetime.strptime(d_str, "%Y-%m-%d").date()
        except:
            return datetime.date.min

    filtered_records = sorted(filtered_records, key=parse_rec_date, reverse=True)

    st.sidebar.markdown(f"**검색 결과: {len(filtered_records)}건**")

    if not filtered_records:
        st.sidebar.info("조회 조건에 해당하는 등록 내역이 없습니다.")
    else:
        # 5개 카드 높이(약 620px)를 기준 삼아 고정하고, 6개부터는 스크롤할 수 있도록 CSS 컨테이너 추가
        # max-height: 620px; overflow-y: auto; 속성 부여
        scrollable_container = st.sidebar.container(height=620)
        
        with scrollable_container:
            # 렌더링 최적화를 위해 최신 50건만 표시
            if len(filtered_records) > 50:
                st.warning(f"검색 결과가 많아 최신 50건만 목록에 표시합니다. (전체 {len(filtered_records)}건 AI 리포트 생성은 정상 동작합니다)")

            for rec in filtered_records[:50]:
                # 카드 형태의 정보 표기
                with st.container(border=True):
                    st.write(f"**연번 {rec['연번']}** | {rec['고객명'] or '고객명 없음'} ({rec['접수 일자']})")
                    st.write(f"*{rec['판매처']} | {rec['대']} > {rec['중']}*")
                    if rec['문의내용']:
                        st.caption(rec['문의내용'][:40] + ("..." if len(rec['문의내용']) > 40 else ""))
                    
                    # 불러오기 버튼 클릭 핸들러
                    if st.button(f"불러오기 📂", key=f"load_{rec['row_idx']}_{rec['연번']}"):
                        # 세션 상태에 복사 데이터 채워넣기
                        st.session_state["edit_mode"] = True
                        st.session_state["edit_row_idx"] = rec["row_idx"]
                        st.session_state["edit_serial"] = rec["연번"]
                        
                        # 폼 바인딩 데이터 설정
                        st.session_state["판매처_선택"] = rec["판매처"] if rec["판매처"] in load_unique_sellers() else "+ 직접 입력 (새로 추가)"
                        if st.session_state["판매처_선택"] == "+ 직접 입력 (새로 추가)":
                            st.session_state["판매처_직접입력"] = rec["판매처"]
                            
                        st.session_state["유형_선택"] = rec["유형"] if rec["유형"] in load_unique_types() else "+ 직접 입력 (새로 추가)"
                        if st.session_state["유형_선택"] == "+ 직접 입력 (새로 추가)":
                            st.session_state["유형_직접입력"] = rec["유형"]
                            
                        st.session_state["고객유형_선택"] = rec["고객유형"] if rec["고객유형"] in load_unique_customer_types() else "+ 직접 입력 (새로 추가)"
                        if st.session_state["고객유형_선택"] == "+ 직접 입력 (새로 추가)":
                            st.session_state["고객유형_직접입력"] = rec["고객유형"]
                        
                        st.session_state["담당자"] = rec["담당자"]
                        try:
                            st.session_state["접수일자"] = datetime.datetime.strptime(rec["접수 일자"][:10], "%Y-%m-%d").date()
                        except:
                            st.session_state["접수일자"] = datetime.date.today()
                            
                        st.session_state["고객명"] = rec["고객명"]
                        st.session_state["고객전화번호"] = rec["고객 전화번호"]
                        st.session_state["대분류"] = rec["대"]
                        st.session_state["중분류"] = rec["중"]
                        st.session_state["소분류"] = rec["소"]
                        st.session_state["문의내용"] = rec["문의내용"]
                        st.session_state["주문여부"] = rec["주문여부"]
                        
                        try:
                            st.session_state["금액"] = int(rec["금액"]) if rec["금액"] else 0
                        except:
                            st.session_state["금액"] = 0
                        
                        st.session_state["첫재주문"] = rec["첫/재주문"]
                        st.session_state["연령대"] = rec["연령대"]
                        st.session_state["성별"] = rec["성별"]
                        st.session_state["인입경로"] = rec["인입경로"]
                        st.session_state["중복여부"] = rec["중복 여부"]
                        
                        # 배송품질 관련 처리
                        has_delivery_info = bool(rec["성함"] or rec["제품명"] or rec["클레임 유형"])
                        st.session_state["show_delivery"] = has_delivery_info
                        
                        st.session_state["재출고여부"] = rec["재출고 여부"]
                        st.session_state["성함"] = rec["성함"]
                        st.session_state["운송장"] = rec["운송장"]
                        st.session_state["제품명"] = rec["제품명"]
                        
                        try:
                            st.session_state["발생수량"] = int(rec["발생수량(EA)"]) if rec["발생수량(EA)"] else 0
                        except:
                            st.session_state["발생수량"] = 0
                            
                        st.session_state["클레임유형"] = rec["클레임 유형"]
                        st.session_state["보상"] = rec["보상"]
                        
                        st.rerun()

    st.sidebar.divider()
    st.sidebar.markdown("### 📊 AI 리포트 기능")
    if st.sidebar.button("🤖 AI 리포트 생성 (현재 조건)", use_container_width=True):
        st.session_state["show_ai_report"] = True
    if st.sidebar.button("📋 입력 폼으로 돌아가기", use_container_width=True):
        st.session_state["show_ai_report"] = False

    if st.session_state.get("show_ai_report"):
        render_ai_report_tab(filtered_records)
        return

    st.title("📋 VOC 데일리 응대 입력 시스템")
    
    # 수정 모드 표시 배너
    if st.session_state.get("edit_mode"):
        st.warning(f"⚠️ 현재 **[수정 모드]** 활성화 상태입니다. (연번: {st.session_state.get('edit_serial')}) 저장 시 구글 시트의 해당 행이 수정(덮어쓰기)됩니다.")
        if st.button("❌ 수정 취소 및 신규 등록으로 전환", use_container_width=True):
            reset_form(keep_shared_values=False)
            st.rerun()
            
    st.caption("고객감동팀 VOC 상담/문의 내용을 입력하고 저장하세요. ※ 굵은 글씨 항목은 필수 입력값입니다.")

    # 폼 리셋 후 페이지 재로드 플래그 처리
    if st.session_state.get("form_reset"):
        st.session_state["form_reset"] = False
        st.rerun()

    # --------------------------------------------------------
    # 섹션 1: 기본 접수 정보
    # --------------------------------------------------------
    st.subheader("1. 기본 접수 정보")

    # 첫 번째 행: 판매처, 유형, 접수 일자
    row1_col1, row1_col2, row1_col3 = st.columns(3)

    with row1_col1:
        # 판매처 목록 로드 및 직접 입력 옵션 제공
        sellers_list = load_unique_sellers()
        sellers_options = sellers_list + ["+ 직접 입력 (새로 추가)"]
        
        default_index = 0
        if "판매처_선택" in st.session_state and st.session_state["판매처_선택"] in sellers_options:
            default_index = sellers_options.index(st.session_state["판매처_선택"])
        
        판매처_선택 = st.selectbox(
            "**판매처** *",
            options=sellers_options,
            index=default_index,
            key="판매처_선택"
        )
        
        if 판매처_선택 == "+ 직접 입력 (새로 추가)":
            판매처 = st.text_input(
                "새 판매처 이름 입력 *",
                key="판매처_직접입력",
                placeholder="예: 현대몰, 11번가"
            )
        else:
            판매처 = 판매처_선택

    with row1_col2:
        # 유형 목록 로드 및 직접 입력 옵션 제공
        types_list = load_unique_types()
        types_options = types_list + ["+ 직접 입력 (새로 추가)"]
        
        default_type_idx = 0
        if "유형_선택" in st.session_state and st.session_state["유형_선택"] in types_options:
            default_type_idx = types_options.index(st.session_state["유형_선택"])
            
        유형_선택 = st.selectbox(
            "**유형** *",
            options=types_options,
            index=default_type_idx,
            key="유형_선택"
        )
        
        if 유형_선택 == "+ 직접 입력 (새로 추가)":
            유형 = st.text_input(
                "새 유형 이름 입력 *",
                key="유형_직접입력",
                placeholder="예: 이메일, 채팅상담"
            )
        else:
            유형 = 유형_선택

    with row1_col3:
        접수일자 = st.date_input(
            "**접수 일자** *",
            value=datetime.date.today(),
            key="접수일자",
        )

    # 두 번째 행: 담당자, 고객유형, 고객명
    row2_col1, row2_col2, row2_col3 = st.columns(3)

    with row2_col1:
        담당자 = st.text_input("**담당자** *", key="담당자",
                               placeholder="담당자 이름")

    with row2_col2:
        # 고객유형 목록 로드 및 직접 입력 옵션 제공
        cust_types_list = load_unique_customer_types()
        cust_types_options = cust_types_list + ["+ 직접 입력 (새로 추가)"]
        
        default_cust_type_idx = 0
        if "고객유형_선택" in st.session_state and st.session_state["고객유형_선택"] in cust_types_options:
            default_cust_type_idx = cust_types_options.index(st.session_state["고객유형_선택"])
            
        고객유형_선택 = st.selectbox(
            "**고객유형** *",
            options=cust_types_options,
            index=default_cust_type_idx,
            key="고객유형_선택"
        )
        
        if 고객유형_선택 == "+ 직접 입력 (새로 추가)":
            고객유형 = st.text_input(
                "새 고객유형 이름 입력 *",
                key="고객유형_직접입력",
                placeholder="예: VIP고객, 블랙컨슈머"
            )
        else:
            고객유형 = 고객유형_선택

    with row2_col3:
        고객명 = st.text_input("고객명", key="고객명", placeholder="홍길동")

    고객전화번호 = st.text_input(
        "고객 전화번호",
        key="고객전화번호",
        placeholder="010-0000-0000",
    )

    st.divider()

    # --------------------------------------------------------
    # 섹션 2: CS 분류 (대 > 중 > 소 종속 드롭다운)
    # --------------------------------------------------------
    st.subheader("2. CS 분류")

    large_options = [""] + get_large_categories()
    
    default_large_idx = 0
    if "대분류" in st.session_state and st.session_state["대분류"] in large_options:
        default_large_idx = large_options.index(st.session_state["대분류"])
        
    대분류 = st.selectbox("**대분류** *", options=large_options, index=default_large_idx, key="대분류")

    middle_options = [""] + get_middle_categories(대분류) if 대분류 else [""]
    
    default_middle_idx = 0
    if "중분류" in st.session_state and st.session_state["중분류"] in middle_options:
        default_middle_idx = middle_options.index(st.session_state["중분류"])
        
    중분류 = st.selectbox("**중분류** *", options=middle_options, index=default_middle_idx, key="중분류")

    small_options_raw = get_small_categories(대분류, 중분류)
    if small_options_raw:
        small_options = [""] + small_options_raw
        
        default_small_idx = 0
        if "소분류" in st.session_state and st.session_state["소분류"] in small_options:
            default_small_idx = small_options.index(st.session_state["소분류"])
            
        소분류 = st.selectbox("**소분류** *", options=small_options, index=default_small_idx, key="소분류")
        소분류_label = "**소분류** *"
    else:
        소분류 = st.selectbox(
            "소분류 (해당 없음)",
            options=[""],
            key="소분류",
            disabled=True,
            help="선택한 중분류에는 소분류가 없습니다.",
        )
        소분류_label = "소분류"

    st.divider()

    # --------------------------------------------------------
    # 섹션 3: 문의내용
    # --------------------------------------------------------
    st.subheader("3. 문의내용")

    문의내용 = st.text_area(
        "**문의내용** *",
        key="문의내용",
        height=150,
        placeholder="고객 문의 내용을 상세히 입력하세요.",
    )

    st.divider()

    # --------------------------------------------------------
    # 섹션 4: 전화주문 정보
    # --------------------------------------------------------
    st.subheader("4. 전화주문 정보")

    col_a, col_b = st.columns([1, 3])
    with col_a:
        주문여부 = st.selectbox(
            "주문여부",
            options=["", "O", "X"],
            key="주문여부",
        )
    with col_b:
        중복여부 = st.selectbox(
            "중복 여부",
            options=["", "중복", "중복아님"],
            key="중복여부",
        )

    # 주문여부 O일 때만 세부 항목 표시
    금액 = ""
    첫재주문 = ""
    연령대 = ""
    성별 = ""
    인입경로 = ""

    if 주문여부 == "O":
        with st.expander("📦 주문 세부 정보 입력", expanded=True):
            col_x, col_y = st.columns(2)
            with col_x:
                금액 = st.number_input(
                    "**금액** *",
                    min_value=0,
                    step=100,
                    key="금액",
                    help="주문 금액 (원)",
                )
                첫재주문 = st.selectbox(
                    "**첫/재주문** *",
                    options=["", "첫주문", "재주문"],
                    key="첫재주문",
                )
                연령대 = st.selectbox(
                    "연령대",
                    options=["", "10대", "20대", "30대", "40대", "50대", "60대 이상"],
                    key="연령대",
                )
            with col_y:
                성별 = st.selectbox(
                    "성별",
                    options=["", "남", "여", "기타"],
                    key="성별",
                )
                인입경로 = st.text_input(
                    "인입경로",
                    key="인입경로",
                    placeholder="예: 네이버, 인스타그램",
                )

    st.divider()

    # --------------------------------------------------------
    # 섹션 5: 오출고 CS율 / 배송품질 정보
    # --------------------------------------------------------
    st.subheader("5. 오출고 CS율 / 배송품질 정보")

    show_delivery = st.checkbox(
        "📦 배송품질 정보 입력",
        key="show_delivery",
        help="배송품질 관련 정보가 있는 경우 체크하세요.",
    )

    재출고여부 = ""
    성함 = ""
    운송장 = ""
    제품명 = ""
    발생수량 = ""
    클레임유형 = ""
    보상 = ""

    if show_delivery:
        with st.expander("🚚 배송품질 세부 정보 입력", expanded=True):
            col_p, col_q = st.columns(2)
            with col_p:
                재출고여부 = st.selectbox(
                    "재출고 여부",
                    options=["", "O", "X"],
                    key="재출고여부",
                )
                성함 = st.text_input("**성함** *", key="성함", placeholder="홍길동")
                운송장 = st.text_input(
                    "운송장",
                    key="운송장",
                    placeholder="운송장 번호",
                )
                제품명 = st.text_input(
                    "**제품명** *",
                    key="제품명",
                    placeholder="예: 레몬500ML",
                )
            with col_q:
                발생수량 = st.number_input(
                    "**발생수량(EA)** *",
                    min_value=0,
                    step=1,
                    key="발생수량",
                )
                클레임유형 = st.selectbox(
                    "**클레임 유형** *",
                    options=[
                        "", "파손", "송장탈착", "지연", "오배송",
                        "분실", "출고누락", "오출고", "내품파손", "기타",
                    ],
                    key="클레임유형",
                )
                보상 = st.text_input(
                    "보상",
                    key="보상",
                    placeholder="예: 레몬500ML_1EA",
                )

    st.divider()

    # --------------------------------------------------------
    # 미리보기
    # --------------------------------------------------------
    with st.expander("👁️ 입력 내용 미리보기", expanded=False):
        preview_data = {
            "판매처": 판매처,
            "담당자": 담당자,
            "유형": 유형,
            "접수 일자": str(접수일자),
            "고객유형": 고객유형,
            "고객명": 고객명,
            "고객 전화번호": 고객전화번호,
            "대": 대분류,
            "중": 중분류,
            "소": 소분류,
            "문의내용": 문의내용[:50] + "..." if len(문의내용) > 50 else 문의내용,
            "주문여부": 주문여부,
            "금액": 금액 if 주문여부 == "O" else "",
            "첫/재주문": 첫재주문 if 주문여부 == "O" else "",
            "연령대": 연령대 if 주문여부 == "O" else "",
            "성별": 성별 if 주문여부 == "O" else "",
            "인입경로": 인입경로 if 주문여부 == "O" else "",
            "중복 여부": 중복여부,
        }
        if show_delivery:
            preview_data.update({
                "재출고 여부": 재출고여부,
                "성함": 성함,
                "운송장": 운송장,
                "제품명": 제품명,
                "발생수량(EA)": 발생수량,
                "클레임 유형": 클레임유형,
                "보상": 보상,
            })

        cols = st.columns(3)
        items = list(preview_data.items())
        chunk = len(items) // 3 + 1
        for i, col in enumerate(cols):
            with col:
                for k, v in items[i * chunk : (i + 1) * chunk]:
                    st.write(f"**{k}**: {v}")

    # --------------------------------------------------------
    # 연속 등록 여부 체크박스 & 저장/수정 버튼
    # --------------------------------------------------------
    keep_values = st.checkbox(
        "🔁 저장 후 공통 입력값 유지 (판매처, 담당자, 유형, 접수일자, 대/중/소분류 등)",
        value=False,
        key="keep_shared_values",
        help="체크 시 저장 완료 후 고객명, 전화번호, 문의내용만 지워지고 공통 분류 정보 등은 그대로 유지되어 연속 입력이 편리해집니다."
    )

    is_edit_mode = st.session_state.get("edit_mode", False)
    btn_label = "💾 수정 완료 (구글 시트 덮어쓰기)" if is_edit_mode else "💾 신규 저장"

    if st.button(btn_label, type="primary", use_container_width=True):

        # 폼 데이터 수집
        form_data = {
            "판매처": 판매처,
            "담당자": 담당자,
            "유형": 유형,
            "접수 일자": str(접수일자),
            "고객유형": 고객유형,
            "고객명": 고객명,
            "고객 전화번호": 고객전화번호,
            "대": 대분류,
            "중": 중분류,
            "소": 소분류,
            "문의내용": 문의내용,
            "주문여부": 주문여부,
            "금액": 금액 if 주문여부 == "O" else "",
            "첫/재주문": 첫재주문 if 주문여부 == "O" else "",
            "연령대": 연령대 if 주문여부 == "O" else "",
            "성별": 성별 if 주문여부 == "O" else "",
            "인입경로": 인입경로 if 주문여부 == "O" else "",
            "중복 여부": 중복여부,
            "재출고 여부": 재출고여부 if show_delivery else "",
            "성함": 성함 if show_delivery else "",
            "운송장": 운송장 if show_delivery else "",
            "제품명": 제품명 if show_delivery else "",
            "발생수량(EA)": 발생수량 if show_delivery else "",
            "클레임 유형": 클레임유형 if show_delivery else "",
            "보상": 보상 if show_delivery else "",
        }

        # 유효성 검사
        errors = validate_form(form_data, show_delivery)
        if errors:
            for err in errors:
                st.error(f"❌ {err}")
        else:
            # 워크시트 연결
            worksheet = get_worksheet()
            if worksheet is None:
                st.error("Google Sheets 연결에 실패했습니다. secrets 설정을 확인하세요.")
            else:
                if is_edit_mode:
                    # 수정 모드: 기존 연번, 월번, NO 유지
                    row_idx = st.session_state["edit_row_idx"]
                    edit_serial = st.session_state["edit_serial"]
                    
                    # 기존 번호들은 사이드바 로드한 것에서 찾거나 재생성 방지
                    # 안전을 위해 구글 시트의 해당 행에서 직접 앞번호 3개를 읽거나 기존 세션 값을 유지
                    # 여기서는 기존 연번 유지 및 구글 시트에 업데이트할 행 데이터 구성
                    
                    # gspread를 통한 행 값 재조회(동기화 유무)가 필요할 수 있으나,
                    # 불러올 당시의 연번을 그대로 넣어줌
                    # 기존의 연번, 월번, NO 유지
                    recent_list = load_all_records()
                    target_rec = next((r for r in recent_list if r["row_idx"] == row_idx), None)
                    
                    if target_rec:
                        cur_serial = target_rec["연번"]
                        cur_monthly = target_rec["월번"]
                        cur_no = target_rec["NO"]
                    else:
                        cur_serial = edit_serial
                        cur_monthly = ""
                        cur_no = edit_serial

                    row_data = [
                        cur_serial,                               # 연번
                        cur_monthly,                              # 월번
                        cur_no,                                   # NO
                        form_data["판매처"],                      # 판매처
                        form_data["담당자"],                      # 담당자
                        form_data["유형"],                        # 유형
                        form_data["접수 일자"],                   # 접수 일자
                        form_data["고객유형"],                    # 고객유형
                        form_data["고객명"],                      # 고객명
                        form_data["고객 전화번호"],               # 고객 전화번호
                        form_data["대"],                          # 대
                        form_data["중"],                          # 중
                        form_data["소"],                          # 소
                        form_data["문의내용"],                    # 문의내용
                        form_data["주문여부"],                    # 주문여부
                        form_data["금액"] if form_data["금액"] != "" else "",  # 금액
                        form_data["첫/재주문"],                   # 첫/재주문
                        form_data["연령대"],                      # 연령대
                        form_data["성별"],                        # 성별
                        form_data["인입경로"],                    # 인입경로
                        form_data["중복 여부"],                   # 중복 여부
                        form_data["재출고 여부"],                 # 재출고 여부
                        form_data["성함"],                        # 성함
                        form_data["운송장"],                      # 운송장
                        form_data["제품명"],                      # 제품명
                        form_data["발생수량(EA)"] if form_data["발생수량(EA)"] != "" else "",  # 발생수량(EA)
                        form_data["클레임 유형"],                 # 클레임 유형
                        form_data["보상"],                        # 보상
                    ]
                    
                    success = update_sheet(worksheet, row_idx, row_data)
                    if success:
                        st.success(f"✅ 연번 {edit_serial}번 항목이 성공적으로 수정되었습니다!")
                        st.balloons()
                        st.cache_data.clear()
                        # 리셋
                        reset_form(keep_shared_values=keep_values)
                else:
                    # 신규 등록 모드
                    next_serial, next_monthly, next_no = get_next_numbers(
                        worksheet, 접수일자
                    )

                    # 저장 데이터 구성 (컬럼 순서 고정)
                    row_data = [
                        next_serial,                              # 연번
                        next_monthly,                             # 월번
                        next_no,                                  # NO
                        form_data["판매처"],                      # 판매처
                        form_data["담당자"],                      # 담당자
                        form_data["유형"],                        # 유형
                        form_data["접수 일자"],                   # 접수 일자
                        form_data["고객유형"],                    # 고객유형
                        form_data["고객명"],                      # 고객명
                        form_data["고객 전화번호"],               # 고객 전화번호
                        form_data["대"],                          # 대
                        form_data["중"],                          # 중
                        form_data["소"],                          # 소
                        form_data["문의내용"],                    # 문의내용
                        form_data["주문여부"],                    # 주문여부
                        form_data["금액"] if form_data["금액"] != "" else "",  # 금액
                        form_data["첫/재주문"],                   # 첫/재주문
                        form_data["연령대"],                      # 연령대
                        form_data["성별"],                        # 성별
                        form_data["인입경로"],                    # 인입경로
                        form_data["중복 여부"],                   # 중복 여부
                        form_data["재출고 여부"],                 # 재출고 여부
                        form_data["성함"],                        # 성함
                        form_data["운송장"],                      # 운송장
                        form_data["제품명"],                      # 제품명
                        form_data["발생수량(EA)"] if form_data["발생수량(EA)"] != "" else "",  # 발생수량(EA)
                        form_data["클레임 유형"],                 # 클레임 유형
                        form_data["보상"],                        # 보상
                    ]

                    # 저장 실행
                    success = append_to_sheet(worksheet, row_data)
                    if success:
                        st.success(
                            f"✅ 신규 저장 완료되었습니다! (연번: {next_serial}, 월번: {next_monthly})"
                        )
                        st.balloons()
                        # 캐시 초기화하여 새로 추가된 판매처 반영
                        st.cache_data.clear()
                        # 폼 초기화
                        reset_form(keep_shared_values=keep_values)


if __name__ == "__main__":
    main()
