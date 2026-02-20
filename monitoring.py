import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time

# 설정값 (Settings)
URL = "https://www.nia.or.kr/site/nia_kor/ex/bbs/List.do?cbIdx=78336"
EXCEL_FILE = r'C:\Users\default.DESKTOP-JNPQLP7\Desktop\업무\게시판 모니터링 인수인계\최신화 ★2026_주요기관 게시판 및 디전_사전협의 모니터링_최신( 함형범 전임 ).xlsx'
SHEET_NAME = '게시판(NIA)'  # 실제 시트명으로 수정 필수

# 엑셀 파일 읽기
def read_excel_data(file_path, sheet_name='게시판(NIA)', position='bottom'):
    try:
        # 엑셀 파일 로드
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # 최상단 혹은 최하단 데이터 추출
        if position == 'top':
            # 첫 번째 행의 '사업명' 컬럼 데이터 가져오기
            target_name = df['사업명'].iloc[0]
        else:
            # 마지막 행의 '사업명' 컬럼 데이터 가져오기
            target_name = df['사업명'].iloc[-1]

        return target_name
    except Exception as e:
        print(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
        return None


# 웹 모니터링 함수 (Web Monitoring Function)
def start_monitoring():

    # 엑셀 데이터 로드
    target_name = read_excel_data(EXCEL_FILE, SHEET_NAME, position='bottom')
    print(f"🎯 찾고 있는 사업명: {target_name}")

    # 브라우저 실행
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')  # 화면 없이 실행하고 싶을 때 주석 제거
    options.add_experimental_option("detach", True)  # 실행 완료 후 브라우저 종료 방지

    # 크롬 드라이버 자동 설치 및 설정
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    driver.get(URL)

    new_post_links = []  # 새로운 게시물들의 a 태그 정보를 담을 리스트
    stop_searching = False

    while True:
        try:
            # 페이지가 로드될 때까지 잠시 대기
            time.sleep(3)

            # 1. board_type01 클래스 내부의 모든 li 요소를 가져옴
            li_elements = driver.find_elements(By.CSS_SELECTOR, ".board_type01 ul li")

            for li in li_elements:
                try:

                    # 2. li 내부의 제목(subject 클래스) 텍스트 추출
                    title_el = li.find_element(By.CSS_SELECTOR, ".subject")
                    title_text = title_el.text.strip()

                    # 3. 엑셀의 target과 비교 (발견 시 즉시 중단)
                    if target_name in title_text:
                        print(f"🏁 기준점 발견: [{title_text}]. 수집을 중단합니다.")
                        stop_searching = True
                        break

                    # 4. 기준점 이전의 항목이라면 a 태그 정보(onclick 등)를 수집
                    a_tag = li.find_element(By.TAG_NAME, "a")
                    onclick_val = a_tag.get_attribute("onclick")

                    if title_text not in [item['title'] for item in new_post_links]:
                        print(f"🆕 새 공고 수집: {title_text}")
                        new_post_links.append({
                            'title': title_text,
                            'onclick': onclick_val  # 요소 대신 스크립트 실행값 저장
                        })

                    # 중복 수집 방지
                    if title_text not in [item['title'] for item in new_post_links]:
                        print(f"🆕 새 공고 수집: {title_text}")
                        new_post_links.append({
                            'title': title_text,
                            'element': a_tag,
                            'onclick': a_tag.get_attribute("onclick")
                        })
                except:
                    continue  # 구조가 다른 li(공지사항 등)는 건너뜁니다.

            if stop_searching:
                break  # 전체 페이지 탐색 종료
            else:
                print("👉 현재 페이지에 없음. 다음 페이지로 이동 시도...")
                try:
                    # 1. 현재 활성화된 페이지 번호 찾기
                    current_page_el = driver.find_element(By.CSS_SELECTOR, ".pageNation a.active")
                    current_page_num = int(current_page_el.text)
                    next_page_num = current_page_num + 1

                    print(f"현재 페이지: {current_page_num} -> 다음 목표 페이지: {next_page_num}")

                    # 2. 다음 숫자를 클릭하거나, 없으면 '다음' 버튼 클릭
                    try:
                        next_page_link = driver.find_element(By.LINK_TEXT, str(next_page_num))
                        next_page_link.click()
                        print(f"👉 {next_page_num}페이지로 이동합니다.")
                    except:
                        print(f"⚠️ {next_page_num}번 링크가 없습니다. '다음 페이지' 버튼을 사용합니다.")
                        next_btn = driver.find_element(By.CSS_SELECTOR, ".pageNation .next a")
                        next_btn.click()

                    time.sleep(3)  # 페이지 로딩 대기

                except Exception as page_error:
                    print(f"❌ 페이지 이동 중 오류가 발생했습니다: {page_error}")
                    break


        except Exception as e:
            print(f"⚠️ 검색 중 오류 발생: {e}")
            break

    return new_post_links


def extract_detail_info(driver):

    def normalize(s):
        return " ".join(str(s).split()).strip()

    try:
        time.sleep(2)  # 페이지 로딩 대기

        # 1. 제목 (Title)
        title = normalize(driver.find_element(By.CLASS_NAME, "tit_area").text)

        # 2. 날짜 (Date)
        reg_date = driver.find_element(By.CSS_SELECTOR, ".write_area .src em").text.strip()

        # 3. 담당자 및 팀명 (Manager and Team)
        writer_elements = driver.find_elements(By.CSS_SELECTOR, ".write_area .writer em")
        manager = writer_elements[0].text.strip() if len(writer_elements) > 0 else "N/A"
        team = writer_elements[1].text.strip() if len(writer_elements) > 1 else "N/A"

        return {
            "title": title,
            "reg_date": reg_date,
            "manager": manager,
            "team": team
        }
    except Exception as e:
        print(f"❌ 상세 정보 추출 실패: {e}")
        return None

links = start_monitoring()
final_data = []
if links:
    # 브라우저 실행
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')  # 화면 없이 실행하고 싶을 때 주석 제거
    options.add_experimental_option("detach", True)  # 실행 완료 후 브라우저 종료 방지

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()
    driver.get(URL)

    # 2. 하나의 드라이버 세션에서 순차적으로 작업 (Work sequentially in one driver session)
    # 여기서는 start_monitoring에서 썼던 driver를 그대로 사용해야 합니다.
    # The driver used in start_monitoring should be reused here.
    for i, item in enumerate(links, 1):
        print(f"🚀 [{i}/{len(links)}] 상세 페이지 분석 중: {item['title'][:20]}...")

        # 상세페이지 진입
        driver.execute_script(item['onclick'])

        # 정보 추출
        details = extract_detail_info(driver)
        if details:
            final_data.append(details)
            print(f"   ✅ 완료: {details['title']} {details['reg_date']} {details['manager']} ({details['team']})")

        # 다시 목록으로 돌아가기 (Go back to the list)
        driver.back()
        time.sleep(2)

    print(f"\n✨ 총 {len(final_data)}건의 상세 정보 수집이 완료되었습니다.")
