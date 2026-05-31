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


@st.cache_data(ttl=600)  # API 호출을 최소화하기 위해 10분 캐싱
def load_unique_sellers():
    """스프레드시트에서 고유한 판매처 목록을 동적으로 로드합니다."""
    worksheet = get_worksheet()
    default_sellers = ["공식몰", "네이버스마트스토어", "쿠팡", "지마켓", "카카오톡 선물하기"]
    if worksheet is None:
        return default_sellers
    try:
        # D열 (4번째 열)에 입력된 판매처 데이터를 모두 가져옵니다.
        col_values = worksheet.col_values(4)
        # 헤더 및 제외할 텍스트 필터링
        exclude_values = ["판매처", "통화내역/온라인 접수내역"]
        sellers = [val.strip() for val in col_values[1:] if val.strip() and val.strip() not in exclude_values]
        # 중복 제거 및 정렬
        unique_sellers = sorted(list(set(sellers)))
        
        # 기본 판매처가 결과 목록에 없으면 추가
        for ds in default_sellers:
            if ds not in unique_sellers:
                unique_sellers.append(ds)
        return sorted(unique_sellers)
    except Exception as e:
        # 에러 발생 시 기본값 반환
        return default_sellers


@st.cache_data(ttl=600)  # API 호출을 최소화하기 위해 10분 캐싱
def load_unique_types():
    """스프레드시트에서 고유한 유형 목록을 동적으로 로드합니다."""
    worksheet = get_worksheet()
    default_types = ["온라인", "유선"]
    if worksheet is None:
        return default_types
    try:
        # F열 (6번째 열)에 입력된 유형 데이터를 모두 가져옵니다.
        col_values = worksheet.col_values(6)
        # 헤더 및 제외할 텍스트 필터링
        exclude_values = ["유형", "통화내역/온라인 접수내역"]
        types = [val.strip() for val in col_values[1:] if val.strip() and val.strip() not in exclude_values]
        unique_types = sorted(list(set(types)))
        
        # 기본 유형이 결과 목록에 없으면 추가
        for dt in default_types:
            if dt not in unique_types:
                unique_types.append(dt)
        return sorted(unique_types)
    except Exception as e:
        # 에러 발생 시 기본값 반환
        return default_types


@st.cache_data(ttl=600)  # API 호출을 최소화하기 위해 10분 캐싱
def load_unique_customer_types():
    """스프레드시트에서 고유한 고객유형 목록을 동적으로 로드합니다."""
    worksheet = get_worksheet()
    default_cust_types = ["일반고객", "강성고객", "단골고객"]
    if worksheet is None:
        return default_cust_types
    try:
        # H열 (8번째 열)에 입력된 고객유형 데이터를 모두 가져옵니다.
        col_values = worksheet.col_values(8)
        # 헤더 및 제외할 텍스트 필터링
        exclude_values = ["고객유형", "통화내역/온라인 접수내역"]
        cust_types = [val.strip() for val in col_values[1:] if val.strip() and val.strip() not in exclude_values]
        unique_cust_types = sorted(list(set(cust_types)))
        
        # 기본 고객유형이 결과 목록에 없으면 추가
        for dct in default_cust_types:
            if dct not in unique_cust_types:
                unique_cust_types.append(dct)
        return sorted(unique_cust_types)
    except Exception as e:
        # 에러 발생 시 기본값 반환
        return default_cust_types




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


# ============================================================
# 폼 초기화
# ============================================================

def reset_form():
    """Session State를 초기화하여 폼을 리셋합니다."""
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
        if key in st.session_state:
            del st.session_state[key]
    st.session_state["form_reset"] = True


# ============================================================
# Streamlit 앱 메인
# ============================================================

def main():
    st.set_page_config(
        page_title="VOC 데일리 응대 입력 시스템",
        page_icon="📋",
        layout="wide",
    )

    st.title("📋 VOC 데일리 응대 입력 시스템")
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
    대분류 = st.selectbox("**대분류** *", options=large_options, key="대분류")

    middle_options = [""] + get_middle_categories(대분류) if 대분류 else [""]
    중분류 = st.selectbox("**중분류** *", options=middle_options, key="중분류")

    small_options_raw = get_small_categories(대분류, 중분류)
    if small_options_raw:
        small_options = [""] + small_options_raw
        소분류 = st.selectbox("**소분류** *", options=small_options, key="소분류")
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
    # 저장 버튼
    # --------------------------------------------------------
    if st.button("💾 저장", type="primary", use_container_width=True):

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
                # 자동 번호 생성
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
                        f"✅ 저장 완료되었습니다! (연번: {next_serial}, 월번: {next_monthly})"
                    )
                    st.balloons()
                    # 캐시 초기화하여 새로 추가된 판매처 반영
                    st.cache_data.clear()
                    # 폼 초기화
                    reset_form()


if __name__ == "__main__":
    main()
