import time
import os
import glob
import re
import datetime

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 다운로드 폴더 지정 (환경 변수 DOWNLOAD_DIR 우선, 없으면 기본 경로 사용)
download_dir = os.environ.get("DOWNLOAD_DIR", r"C:\Users\ohpy0\OneDrive\바탕 화면\주문변환")

# 날짜 지정 (초기값: 오늘)
today = datetime.date.today()
start_date = today.strftime('%Y-%m-%d')
end_date = today.strftime('%Y-%m-%d')

# Chrome WebDriver 옵션 설정 (비밀번호 자동저장/보안 경고 팝업 비활성화)
options = Options()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True,
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False,
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--disable-extensions")
options.add_argument("--disable-extensions-file-access-check")
options.add_argument("--disable-extensions-http-throttling")
options.add_argument("--disable-infobars")
options.add_argument("--enable-automation")
options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)

# 1. 온채널 로그인 페이지 이동
driver.get("https://www.onch3.co.kr/login/login_web.php")
time.sleep(2)

# 2. 로그인 자동 입력 및 버튼 클릭
onch_id = os.environ.get("ONCH_ID", "")
onch_pw = os.environ.get("ONCH_PW", "")
id_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body > div.container > form > div:nth-child(2) > input")))
id_input.clear()
id_input.send_keys(onch_id)
pw_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body > div.container > form > div:nth-child(3) > input")))
pw_input.clear()
pw_input.send_keys(onch_pw)

login_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.container > form > button")))
login_btn.click()
time.sleep(2)

# 3. 주문조회 페이지 이동
driver.get("https://www.onch3.co.kr/supplier/orders.php?state=preparing&detailState=preparing")
time.sleep(2)

# 4. 주문내역 다운로드 버튼 클릭
download_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#kt_app_wrapper > div.contentWrap > div.container.p-3 > div.contentWrap > div.card > div.card-header.min-h-auto.p-6.border-bottom-0.d-flex.justify-content-between.align-items-center > div > button.btn.btn-excel.me-2")))
download_btn.click()
time.sleep(2)

# 5. 날짜 입력 (모달 내)
start_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#downExcelOrderListModal > div > div > div.modal-body.p-4 > div.row.align-items-center.mb-4 > div.col > div > input:nth-child(1)")))
start_input.clear()
start_input.send_keys(start_date)
end_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#downExcelOrderListModal > div > div > div.modal-body.p-4 > div.row.align-items-center.mb-4 > div.col > div > input:nth-child(3)")))
end_input.clear()
end_input.send_keys(end_date)

# 6. 체크박스 선택
checkbox = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#agreement-order-down-check")))
if not checkbox.is_selected():
    checkbox.click()
time.sleep(0.5)

# 7. 다운로드 버튼 클릭
download_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#btn-order-excel-down")))
download_btn.click()
time.sleep(2)

# 8. 확인(alert) 버튼 클릭 (다운로드 확인)
try:
    alert = driver.switch_to.alert
    alert.accept()
    time.sleep(2)
except Exception:
    pass


# 엑셀 변환 헬퍼 함수

def try_extract_model_code(s):
    if not isinstance(s, str):
        return ""
    s = s.strip()
    if not s:
        return ""
    last_tok = re.split(r"\s+", s)[-1].strip(".,;:/\\()[]{}!?")
    if re.match(r".*[A-Za-z]+.*-.*\d+", last_tok):
        return last_tok
    return ""


def model_code_or_original(s):
    code = try_extract_model_code(s)
    return code if code else s


def remove_trailing_code_or_long_number(s):
    if not isinstance(s, str):
        return ""
    s = s.strip()
    if not s:
        return ""
    parts = s.split()
    last_tok = parts[-1].strip(".,;:/\\()[]{}!?")
    if re.match(r".*[A-Za-z]+.*-.*\d+", last_tok) or (last_tok.isdigit() and len(last_tok) >= 5):
        return " ".join(parts[:-1])
    return s


def parse_option_for_style(k):
    k = str(k).strip()
    color, size = "", ""
    if "-" in k:
        c, s = k.split("-", 1)
        color = f"색상:{c.strip()}"
        size = f"사이즈:{s.strip()}"
    elif "(" in k and k.endswith(")"):
        color = f"색상:{k}"
    elif k == "기본옵션":
        color = f"색상:{k}"
    elif k:
        color = f"색상:{k}"
    return color, size


def parse_option_for_cj(k):
    k = str(k).strip()
    color, size = "", ""
    if "-" in k:
        color, size = [x.strip() for x in k.split("-", 1)]
    elif k:
        color = k
    return color, size


def transform_for_styledome(df):
    columns = [
        "판매처", "업체주문번호", "수취인명", "수취인전화번호", "수취인핸드폰번호",
        "우편번호", "배송지주소", "상품코드", "옵션1", "옵션2", "옵션3",
        "주문수량", "배송메모",
    ]
    out = []
    for _, row in df.iterrows():
        color, size = parse_option_for_style(row.get('K', ''))
        out.append([
            row.get('B', ''),                          # 판매처
            row.get('P', ''),                          # 업체주문번호
            row.get('E', ''),                          # 수취인명
            row.get('F', ''),                          # 수취인전화번호
            row.get('G', ''),                          # 수취인핸드폰번호
            row.get('H', ''),                          # 우편번호
            row.get('I', ''),                          # 배송지주소
            model_code_or_original(row.get('J', '')),  # 상품코드
            color,                                     # 옵션1
            size,                                      # 옵션2
            row.get('L', ''),                          # 옵션3
            row.get('S', ''),                          # 주문수량
            row.get('T', ''),                          # 배송메모
        ])
    return pd.DataFrame(out, columns=columns)


def transform_for_frienddome(df):
    columns = [
        "판매 상품명", "상품번호", "제품코드", "옵션", "수량",
        "받는사람", "받는사람 전화번호", "받는사람 핸드폰번호", "우편번호",
        "받는사람 주소", "배송메세지", "자체주문번호",
    ]
    out = []
    for _, row in df.iterrows():
        out.append([
            remove_trailing_code_or_long_number(row.get('J', '')),  # 판매 상품명
            row.get('W', ''),   # 상품번호
            row.get('K', ''),   # 제품코드
            row.get('L', ''),   # 옵션
            row.get('E', ''),   # 수량
            row.get('F', ''),   # 받는사람
            row.get('G', ''),   # 받는사람 전화번호
            row.get('H', ''),   # 받는사람 핸드폰번호
            row.get('I', ''),   # 우편번호
            row.get('S', ''),   # 받는사람 주소
            row.get('T', ''),   # 배송메세지
            row.get('P', ''),   # 자체주문번호
        ])
    return pd.DataFrame(out, columns=columns)


def transform_for_bagseller(df):
    header1 = [
        "가방쟁이 상품코드", "수량", "배송방식", "받는사람 이름", "받는사람 전화번호",
        "받는사람 휴대폰", "우편번호", "받는사람 주소", "옵션1", "옵션2",
        "배송요청사항", "판매자 메모", "원장주문코드",
    ]
    header2 = ["1.필수 입력 항목입니다."] * 13
    header3 = ["주문은 4열부터!"] + [None] * 12

    out = [header1, header2, header3]
    for _, row in df.iterrows():
        out.append([
            row.get('W', ''),   # 가방쟁이 상품코드
            row.get('L', ''),   # 수량
            row.get('R', ''),   # 배송방식
            row.get('E', ''),   # 받는사람 이름
            row.get('F', ''),   # 받는사람 전화번호
            row.get('G', ''),   # 받는사람 휴대폰
            row.get('H', ''),   # 우편번호
            row.get('I', ''),   # 받는사람 주소
            row.get('K', ''),   # 옵션1
            row.get('M', ''),   # 옵션2
            row.get('N', ''),   # 배송요청사항
            row.get('O', ''),   # 판매자 메모
            row.get('P', ''),   # 원장주문코드
        ])
    return pd.DataFrame(out)


def transform_for_cj(df):
    columns = [
        "예약구분", "집하예정일", "받는분성명", "받는분전화번호", "받는분기타연락",
        "받는분우편번호", "받는분주소(전체, 분할)", "운송장번호", "고객주문번호",
        "품목명", "박스수량", "박스타입", "기본운임", "배송메세지1", "배송메세지2",
        "품목명2", "운임구분",
    ]
    out = []
    for _, row in df.iterrows():
        nameJ = row.get('J', '')
        optK = row.get('K', '')
        qtyL = row.get('L', '')
        c, s = parse_option_for_cj(optK)
        opt_text = " ".join([c, s]).strip()
        p_text = f"{nameJ} {opt_text}".strip() if opt_text else nameJ
        if qtyL:
            p_text = f"{p_text} {qtyL}개"
        out.append([
            row.get('A', ''),   # 예약구분
            row.get('B', ''),   # 집하예정일
            row.get('E', ''),   # 받는분성명
            row.get('F', ''),   # 받는분전화번호
            row.get('G', ''),   # 받는분기타연락
            row.get('H', ''),   # 받는분우편번호
            row.get('I', ''),   # 받는분주소(전체, 분할)
            "",                 # 운송장번호 (CJ에서 발급)
            row.get('P', ''),   # 고객주문번호
            p_text,             # 품목명 (상품명+옵션+수량)
            row.get('M', ''),   # 박스수량
            row.get('N', ''),   # 박스타입
            row.get('O', ''),   # 기본운임
            row.get('T', ''),   # 배송메세지1
            row.get('Q', ''),   # 배송메세지2
            row.get('R', ''),   # 품목명 (두 번째)
            row.get('S', ''),   # 운임구분
        ])
    return pd.DataFrame(out, columns=columns)


# 10. 엑셀 파일 읽기 및 변환
output_dir = download_dir
os.makedirs(output_dir, exist_ok=True)
order_files = glob.glob(os.path.join(download_dir, "order_*.xlsx"))
if not order_files:
    raise FileNotFoundError("order_*.xlsx 파일이 다운로드 폴더에 없습니다.")
src_file = max(order_files, key=os.path.getctime)  # 가장 최근 파일
df = pd.read_excel(src_file, dtype=str)
df = df.fillna("")

# 시트별 변환 및 저장
transform_for_styledome(df).to_excel(os.path.join(output_dir, "스타일도매_주문.xlsx"), index=False)
transform_for_frienddome(df).to_excel(os.path.join(output_dir, "친구도매_주문.xlsx"), index=False)
transform_for_bagseller(df).to_excel(os.path.join(output_dir, "가방쟁이_주문.xlsx"), index=False)
transform_for_cj(df).to_excel(os.path.join(output_dir, "CJ등록_주문.xlsx"), index=False)

print("주문 데이터 변환 및 저장 완료! (경로: {})".format(output_dir))

# 브라우저 종료
driver.quit()
