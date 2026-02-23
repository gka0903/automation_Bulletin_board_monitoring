AGENCY_CONFIG = {
    "NIA": {
        "menu_path": ["알림마당", "입찰공고"],
        "url": "https://www.nia.or.kr/site/nia_kor/ex/bbs/List.do?cbIdx=78336",
        "sheet_name": "게시판(NIA)",
        "is_spa": False,
        "selectors": {
            "list_item": ".board_type01 ul li",
            "title": ".subject",
            "next_page": ".pageNation .next a"
        }
    },
    "NIPA": {
        "menu_path": ["알림마당", "입찰공고"],
        "url": "NIPA_URL_입력",
        "sheet_name": "게시판(NIPA)",
        "is_spa": True,
        "selectors": {
            "list_item": "NIPA_목록_선택자",
            "title": "NIPA_제목_선택자",
            "next_page": "NIPA_다음버튼_선택자"
        }
    },
    "KISA": {
        "menu_path": ["알림마당", "입찰공고"],
        "url": "NIPA_URL_입력",
        "sheet_name": "https://www.kisa.or.kr/403",
        "is_spa": False,
        "selectors": {
            "list_item": "NIPA_목록_선택자",
            "title": "NIPA_제목_선택자",
            "next_page": "NIPA_다음버튼_선택자"
        }
    },
    "과기부": {
        "menu_path": ["알림", "사업공고"],
        "url": "https://www.msit.go.kr/bbs/list.do?sCode=user&mPid=121&mId=311",
        "sheet_name": "게시판(과기부)",
        "is_spa": False,  # 싱글 페이지 구조 플래그 활성화
        "selectors": {
            "list_item": "과기부_목록_선택자",
            "title": "과기부_제목_선택자",
            "next_page": "과기부_다음버튼_선택자"
        }
    }
}