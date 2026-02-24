AGENCY_CONFIG = {
    "NIA": {
        "url": "https://www.nia.or.kr/site/nia_kor/ex/bbs/List.do?cbIdx=78336",
        "sheet_name": "게시판(NIA)",
        "link_type": "onclick", # 스크립트 실행은 'onclick', 일반 주소는 'href' /
        "selectors": {
            "list_item": ".board_type01 ul li",
            "title": ".subject",
            "link_tag": "a",
            "active_page": ".pageNation a.active",
            "next_page": ".pageNation .next a",
            "detail_title": ".tit_area",
            "detail_date": ".write_area .src em",
            "detail_writer": ".write_area .writer em"
        },
    },
    "KISA": {
        "url": "https://www.kisa.or.kr/403",
        "sheet_name": "게시판(KISA)",
        "link_type": "href",
        "selectors": {
            "list_item": ".tbl_board tbody tr",
            "title": ".sbj",
            "link_tag": ".sbj a",
            "active_page": ".pagination a.on",
            "next_page": ".pagination a.next",
            "detail_title": ".board_detail_info h2",
            "detail_date": "//dt[contains(text(), '등록일')]/following-sibling::dd", # XPath 사용 / Use XPath
            "detail_team": "//dt[contains(text(), '담당부서')]/following-sibling::dd", # 분리된 부서 정보 / Separated team info
            "detail_manager": "" # 이름이 없으므로 비워둠 / Leave empty as there is no name
        }
    },
    "과기부": {
        "url": "https://www.msit.go.kr/bbs/list.do?sCode=user&mPid=121&mId=311",
        "sheet_name": "게시판(과기부)",
        "link_type": "onclick",
        "selectors": {
            # thead 클래스를 가진 항목은 제외하고 실제 게시글 줄만 선택합니다.
            "list_item": ".board_list .toggle:not(.thead)",
            "title": ".txt .title",
            "link_tag": "a",
            "active_page": ".page-links a[aria-current='page']",
            "next_page": "a.page-navi.next",
            "detail_title": ".view_head h2",
            "detail_date": "//dt[contains(text(), '작성일')]/following-sibling::dd",
            "detail_team": "//dt[contains(text(), '부서')]/following-sibling::dd",
            "detail_manager": "//dt[contains(text(), '담당자')]/following-sibling::dd"
        }
    },
    "NIPA": {
        "menu_path": ["알림마당", "입찰공고"],
        "url": "https://www.nipa.kr/home/2-3",
        "sheet_name": "게시판(NIPA)",
        "link_type": "href",
        "selectors": {
            "list_item": ".bdWrap table tbody tr",
            "title": ".tl a",
            "link_tag": ".tl a",
            "active_page": ".pagination li.active span",
            "next_page": ".pagination a[title*='다음']",
            "detail_title": ".tbTit", # 제목 선택자
            "detail_date": ".tbInfoCon .infoDt",  # 날짜 선택자
            "detail_team": "",  # 팀명은 제공되지 않으므로 비워둠
            "detail_manager": ".tbInfoCon .infoNm" # 담당자(작성자) 선택자
        }
    }
}