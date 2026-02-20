import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from copy import copy
from datetime import datetime
import os
import time
import re

# 설정값
URL = "https://www.nia.or.kr/site/nia_kor/ex/bbs/List.do?cbIdx=78336"
EXCEL_FILE = r'C:\Users\default.DESKTOP-JNPQLP7\Desktop\업무\게시판 모니터링 인수인계\최신화 ★2026_주요기관 게시판 및 디전_사전협의 모니터링_최신( 함형범 전임 ).xlsx'
SHEET_NAME = '게시판(NIA)'


def read_excel_data(file_path, sheet_name='게시판(NIA)'):
    """엑셀에서 마지막에 수집된 사업명을 가져옵니다."""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df['사업명'].iloc[-1]
    except Exception as e:
        print(f"❌ 엑셀 읽기 오류: {e}")
        return None


def start_monitoring(driver, target_name):
    """게시판을 탐색하며 새로운 공고의 스크립트를 수집합니다."""
    driver.get(URL)
    new_post_links = []

    while True:
        time.sleep(3)  # 페이지 로드 대기

        try:
            li_elements = driver.find_elements(By.CSS_SELECTOR, ".board_type01 ul li")
            stop_searching = False

            for li in li_elements:
                try:
                    title_text = li.find_element(By.CSS_SELECTOR, ".subject").text.strip()

                    if target_name in title_text:
                        print(f"🏁 기준점 발견 [{title_text}]. 탐색 중단.")
                        stop_searching = True
                        break

                    a_tag = li.find_element(By.TAG_NAME, "a")
                    onclick_val = a_tag.get_attribute("onclick")

                    # 중복 수집 방지 및 리스트 추가
                    if not any(item['title'] == title_text for item in new_post_links):
                        print(f"🆕 새 공고 수집: {title_text}")
                        new_post_links.append({
                            'title': title_text,
                            'onclick': onclick_val
                        })
                except:
                    continue  # 공지사항 등 구조가 다른 요소 건너뛰기

            if stop_searching:
                break

            # 다음 페이지 이동
            print("👉 현재 페이지에 없음. 다음 페이지로 이동...")
            current_page_num = int(driver.find_element(By.CSS_SELECTOR, ".pageNation a.active").text)
            next_page_num = current_page_num + 1

            try:
                driver.find_element(By.LINK_TEXT, str(next_page_num)).click()
            except:
                driver.find_element(By.CSS_SELECTOR, ".pageNation .next a").click()

        except Exception as e:
            print(f"⚠️ 검색 중 오류 발생: {e}")
            break

    return new_post_links


def extract_detail_info(driver):
    """상세 페이지에서 필요한 정보를 추출합니다."""
    try:
        time.sleep(2)
        raw_title = " ".join(driver.find_element(By.CLASS_NAME, "tit_area").text.split()).strip()
        title = re.sub(r'\[.*?\]', '', raw_title).strip()

        reg_date = driver.find_element(By.CSS_SELECTOR, ".write_area .src em").text.strip().replace('.', '-')

        writer_elements = driver.find_elements(By.CSS_SELECTOR, ".write_area .writer em")
        manager = writer_elements[0].text.strip() if len(writer_elements) > 0 else "N/A"
        team = writer_elements[1].text.strip() if len(writer_elements) > 1 else "N/A"

        return {"title": title, "reg_date": reg_date, "manager": manager, "team": team}
    except Exception as e:
        print(f"❌ 상세 정보 추출 실패: {e}")
        return None


def update_excel_results(file_path, sheet_name, data_list):
    """수집된 데이터를 엑셀에 업데이트하고 서식을 복사합니다."""
    if not data_list:
        return

    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    today = datetime.now().date()

    actual_last_row = ws.max_row
    while actual_last_row > 1 and not ws.cell(row=actual_last_row, column=6).value:
        actual_last_row -= 1

    print(f"📊 현재 실제 데이터 마지막 행: {actual_last_row}")

    for data in reversed(data_list):
        new_row = actual_last_row + 1

        ws.cell(row=new_row, column=1).value = ws.cell(row=actual_last_row, column=1).value
        ws.cell(row=new_row, column=2).value = today
        ws.cell(row=new_row, column=2).number_format = 'yyyy-mm-dd'
        ws.cell(row=new_row, column=3).value = ws.cell(row=actual_last_row, column=3).value
        ws.cell(row=new_row, column=4).value = data['team']
        ws.cell(row=new_row, column=5).value = data['manager']
        ws.cell(row=new_row, column=6).value = data['title']
        ws.cell(row=new_row, column=8).value = data['reg_date']

        for col in range(1, 19):
            source_cell = ws.cell(row=actual_last_row, column=col)
            new_cell = ws.cell(row=new_row, column=col)
            if source_cell.has_style:
                new_cell.font = copy(source_cell.font)
                new_cell.border = copy(source_cell.border)
                new_cell.fill = copy(source_cell.fill)
                new_cell.number_format = copy(source_cell.number_format)
                new_cell.alignment = copy(source_cell.alignment)

        actual_last_row += 1

    wb.save(file_path)
    print(f"💾 엑셀 저장 완료: {len(data_list)}건 추가됨.")


def main():
    target_name = read_excel_data(EXCEL_FILE, SHEET_NAME)
    if not target_name:
        return

    print(f"🎯 찾고 있는 사업명: {target_name}")

    # 1개의 브라우저만 실행하여 재사용
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.maximize_window()

    links = start_monitoring(driver, target_name)
    final_data = []

    if links:
        main_window = driver.current_window_handle
        print(f"🔎 총 {len(links)}개의 신규 공고를 새 탭으로 열어 분석합니다.")

        for i, item in enumerate(links, 1):
            print(f"🚀 [{i}/{len(links)}] 상세 페이지 여는 중: {item['title'][:20]}...")
            try:
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[-1])
                driver.get(URL)
                time.sleep(2)

                driver.execute_script(item['onclick'])
                details = extract_detail_info(driver)

                if details:
                    final_data.append(details)
                    print(f"   ✅ 완료: {details['title']}")

                driver.switch_to.window(main_window)
                time.sleep(1)
            except Exception as e:
                print(f"   ❌ 오류 발생: {e}")
                driver.switch_to.window(main_window)

        update_excel_results(EXCEL_FILE, SHEET_NAME, final_data)
        print(f"\n✨ 총 {len(final_data)}건 수집 완료.")
        print("📢 상세 페이지 탭이 유지되어 있습니다. 수동 작업을 진행하세요.")

    else:
        print("🤷‍♂️ 새로운 공고가 없습니다.")
        driver.quit()  # 새 공고가 없으면 브라우저 종료

    os.startfile(EXCEL_FILE)


if __name__ == "__main__":
    main()