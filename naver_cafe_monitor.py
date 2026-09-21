"""
네이버 카페 모니터링 스크립트 (로그인 없음 - 공개 카페용)
- 엑셀의 카페 URL을 순회하며 글쓴이(ID), 조회수, 댓글 수 수집
- 삭제/권한없음 등은 비고 컬럼에 기록

Requirements:
    pip install selenium webdriver-manager openpyxl pandas
"""

import time
import re
import os
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

DELAY = 2   # 페이지 간 대기 시간(초)



def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def select_excel_file():
    """탐색기에서 엑셀 파일 선택"""
    root = tk.Tk()
    root.withdraw()  # tkinter 기본 창 숨김
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title="엑셀 파일 선택",
        filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]
    )
    root.destroy()
    return file_path


def show_progress_popup(total):
    """추출중 팝업 - 점 애니메이션 + 진행 카운트"""
    import threading

    root = tk.Tk()
    root.title("네이버 카페 데이터 추출")
    root.geometry("340x110")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    root.update_idletasks()
    x = (root.winfo_screenwidth() - 340) // 2
    y = (root.winfo_screenheight() - 110) // 2
    root.geometry(f"340x110+{x}+{y}")

    label_main = tk.Label(root, text="네이버 카페 데이터 추출 중", font=("맑은 고딕", 12, "bold"))
    label_main.pack(pady=(12, 4))
    label_status = tk.Label(root, text="Chrome 드라이버 초기화 중...", font=("맑은 고딕", 10), fg="#333333")
    label_status.pack()
    label_progress = tk.Label(root, text=f"[0 / {total}]", font=("맑은 고딕", 11), fg="#555555")
    label_progress.pack(pady=(4, 0))

    dots = [0]
    current = [0]
    running = [True]
    status = ["Chrome 드라이버 초기화 중..."]

    def tick():
        if not running[0]:
            return
        label_progress.config(text=f"[{current[0]} / {total}]")
        if status[0]:
            label_status.config(text=status[0])
        root.update()
        root.after(200, tick)

    root.after(0, tick)
    root._running = running
    root._current = current
    root._status = status
    return root


def select_save_path(default_name):
    """저장 경로 선택 탐색기"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    save_path = filedialog.asksaveasfilename(
        title="결과 파일 저장",
        initialfile=default_name,
        defaultextension=".xlsx",
        filetypes=[("Excel 파일", "*.xlsx")]
    )
    root.destroy()
    return save_path


def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    driver.set_page_load_timeout(15)  # 15초 안에 안 뜨면 강제 중단
    return driver


def extract_post_info(driver, url):
    result = {"author": "", "views": "", "comments": "", "likes": "", "note": "", "post_date": "", "cafe_name": "", "board_name": "", "title": ""}

    # 카페 URL 로드
    MAX_RETRY = 3
    for attempt in range(1, MAX_RETRY + 1):
        try:
            log(f"  접속 중... (시도 {attempt}/{MAX_RETRY})")
            try:
                driver.execute_script("window.stop();")
            except Exception:
                pass
            driver.get(url)
            # 페이지 로드 직후 alert 체크
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                log(f"  🗑️  alert 감지: {alert_text}")
                if "삭제" in alert_text or "존재하지 않" in alert_text:
                    result["note"] = "삭제된 게시글"
                elif "로그인" in alert_text:
                    result["note"] = "로그인 필요"
                else:
                    result["note"] = f"접근 불가"
                return result
            except Exception:
                pass  # alert 없으면 정상 진행
            break
        except TimeoutException:
            log(f"  ⚠️  로딩 타임아웃 ({attempt}/{MAX_RETRY}) - 재시도...")
            driver.execute_script("window.stop();")
            if attempt == MAX_RETRY:
                log("  ❌ 재시도 횟수 초과")
                result["note"] = "로딩 타임아웃 (재시도 초과)"
                return result
            time.sleep(2)

    try:
        time.sleep(2)

        # alert 팝업 처리 (삭제된 게시글 등)
        try:
            alert = driver.switch_to.alert
            alert_text = alert.text
            alert.accept()
            log(f"  🗑️  alert 감지: {alert_text}")
            if "삭제" in alert_text or "존재하지 않" in alert_text:
                result["note"] = "삭제된 게시글"
            elif "로그인" in alert_text:
                result["note"] = "로그인 필요"
            else:
                result["note"] = "접근 불가"
            return result
        except Exception:
            pass  # alert 없으면 그냥 진행

        # 로그인 페이지로 리다이렉트 됐는지 확인
        current = driver.current_url
        log(f"  [DEBUG] 리다이렉트 URL: {current}")
        if "nidlogin" in current or "nid.naver.com" in current or "login" in current:
            log(f"  🔒 로그인 페이지로 리다이렉트됨 → 로그인 필요")
            result["note"] = "로그인 필요"
            return result

        # 카페명 추출 (외부 페이지 h1.d-none - iframe 진입 전)
        log("  카페명 추출 중...")
        try:
            el = driver.find_element(By.CSS_SELECTOR, "h1.d-none")
            result["cafe_name"] = el.text.strip()
            log(f"  카페명: {result['cafe_name']}")
        except NoSuchElementException:
            log("  ⚠️  카페명 추출 실패")

        # iframe(cafe_main) 안으로 전환
        log("  iframe 진입 중...")
        try:
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "cafe_main"))
            )
            log("  iframe 진입 완료")
        except TimeoutException:
            log("  ⚠️  iframe 없음 - 로그인 필요 가능성 확인...")
            # iframe 자체가 없으면 로그인 필요 카페일 수 있음
            outer_text = driver.find_element(By.TAG_NAME, "body").text
            if "로그인" in outer_text and "카페" in outer_text:
                log("  🔒 로그인 필요 감지")
                result["note"] = "로그인 필요"
                return result

        # 페이지 소스에서 로그인 필요 여부 확인 (iframe 진입 전)
        page_src_outer = driver.page_source
        if "nidlogin" in page_src_outer and "cafe_main" not in page_src_outer:
            log("  🔒 로그인 필요 페이지 감지")
            result["note"] = "로그인 필요"
            return result

        # Vue.js 렌더링 대기
        log("  렌더링 대기 중...")
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".article_info, button.nickname, .ArticleTool"))
            )
            log("  렌더링 완료")
        except TimeoutException:
            log("  ⚠️  렌더링 대기 시간 초과")

        # 디버그: 현재 URL + HTML 파일 저장
        log(f"  [DEBUG] iframe 내부 URL: {driver.current_url}")
        with open("debug_iframe.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        log("  [DEBUG] iframe 소스 → debug_iframe.html 저장")

        page_src = driver.page_source
        body_text = driver.find_element(By.TAG_NAME, "body").text

        # ── 삭제 / 권한 없음 감지 ────────────────────────────
        deleted_kw = ["삭제된 게시글입니다", "삭제된 글입니다", "존재하지 않는 게시글", "없는 게시물입니다"]
        access_kw  = ["접근 권한이 없습니다", "열람 권한이 없습니다", "멤버 공개 게시글", "이 카페의 회원만", "로그인 후 이용하실 수 있습니다", "로그인이 필요한 서비스", "로그인하세요", "로그인이 필요합니다", "카페 가입 후 이용", "로그인 후 이용"]

        for kw in deleted_kw:
            if kw in body_text:
                log(f"  🗑️  삭제된 게시글: '{kw}'")
                result["note"] = "삭제된 게시글"
                driver.switch_to.default_content()
                return result

        for kw in access_kw:
            if kw in body_text:
                note = "로그인 필요" if "로그인" in kw else "권한 없음"
                log(f"  🔒 {note}: '{kw}'")
                result["note"] = note
                driver.switch_to.default_content()
                return result

        log("  정상 게시글 - 데이터 추출 시작")

        # ── 게시판명 ─────────────────────────────────────────
        log("  게시판명 추출 중...")
        try:
            el = driver.find_element(By.CSS_SELECTOR, "a.link_board")
            board_text = driver.execute_script("return arguments[0].childNodes[0].textContent;", el)
            result["board_name"] = board_text.strip()
            log(f"  게시판명: {result['board_name']}")
        except NoSuchElementException:
            log("  ⚠️  게시판명 추출 실패")

        # ── 게시글 제목 ──────────────────────────────────────
        log("  게시글 제목 추출 중...")
        try:
            el = driver.find_element(By.CSS_SELECTOR, "h3.title_text")
            result["title"] = el.text.strip()
            log(f"  게시글 제목: {result['title']}")
        except NoSuchElementException:
            log("  ⚠️  게시글 제목 추출 실패")

        # ── 글쓴이 ───────────────────────────────────────────
        log("  글쓴이 추출 중...")
        try:
            el = driver.find_element(By.CSS_SELECTOR, "button.nickname")
            result["author"] = el.text.strip()
            log(f"  글쓴이: {result['author']}")
        except NoSuchElementException:
            log("  ⚠️  글쓴이 추출 실패")

        # ── 게시일시 ─────────────────────────────────────────
        log("  게시일시 추출 중...")
        try:
            el = driver.find_element(By.CSS_SELECTOR, ".article_info .date")
            result["post_date"] = el.text.strip()
            log(f"  게시일시: {result['post_date']}")
        except NoSuchElementException:
            result["post_date"] = ""
            log("  ⚠️  게시일시 추출 실패")

        # ── 조회수 ───────────────────────────────────────────
        log("  조회수 추출 중...")
        try:
            el = driver.find_element(By.CSS_SELECTOR, ".article_info .count")
            m = re.search(r"(\d[\d,]*)", el.text)
            if m:
                result["views"] = int(m.group(1).replace(",", ""))
                log(f"  조회수: {result['views']}")
        except NoSuchElementException:
            log("  ⚠️  조회수 추출 실패")

        # ── 댓글 수 ──────────────────────────────────────────
        log("  댓글 수 추출 중...")
        try:
            el = driver.find_element(By.CSS_SELECTOR, ".box_left .button_comment strong.num, .ArticleTool .button_comment strong.num")
            txt = el.text.strip().replace(",", "")
            if txt.isdigit():
                result["comments"] = int(txt)
                log(f"  댓글 수: {result['comments']}")
        except NoSuchElementException:
            log("  ⚠️  댓글 수 추출 실패")

        # ── 좋아요 ───────────────────────────────────────────
        log("  좋아요 추출 중...")
        try:
            el = driver.find_element(By.CSS_SELECTOR, ".u_likeit_list_btn em._count")
            txt = el.text.strip().replace(",", "")
            result["likes"] = int(txt) if txt.isdigit() else 0
            log(f"  좋아요: {result['likes']}")
        except NoSuchElementException:
            result["likes"] = 0
            log("  ⚠️  좋아요 추출 실패")

        # 아무것도 못 가져왔으면 로그인 필요로 처리
        if not result["author"] and result["views"] == "" and result["comments"] == "" and not result["note"]:
            log("  🔒 데이터 추출 실패 → 로그인 필요로 처리")
            result["note"] = "로그인 필요"

        driver.switch_to.default_content()

    except Exception as e:
        driver.switch_to.default_content()
        err = str(e)
        # alert 메시지 정제
        if "Alert Text:" in err:
            import re as _re
            m = _re.search(r"Alert Text: (.+?)(\n|Message:)", err)
            alert_msg = m.group(1).strip() if m else "alert 오류"
            if "삭제" in alert_msg or "존재하지 않" in alert_msg:
                result["note"] = "삭제된 게시글"
            elif "로그인" in alert_msg:
                result["note"] = "로그인 필요"
            else:
                result["note"] = f"접근 불가: {alert_msg[:30]}"
        elif "timeout" in err.lower():
            result["note"] = "로딩 타임아웃"
        else:
            result["note"] = f"오류: {err[:40]}"
        log(f"  ❌ 예외: {result['note']}")

    return result


def run():
    log("=" * 50)
    log("네이버 카페 모니터링 시작")
    log("=" * 50)

    excel_path = select_excel_file()
    if not excel_path:
        log("파일을 선택하지 않았습니다. 종료합니다.")
        return

    folder = os.path.dirname(excel_path)
    base   = os.path.splitext(os.path.basename(excel_path))[0]
    today  = datetime.now().strftime("%Y%m%d")
    output_path = os.path.join(folder, f"{base}_{today}_result.xlsx")

    log(f"파일: {excel_path}")

    df = pd.read_excel(excel_path, dtype=str)
    df.columns = df.columns.str.strip()

    for col in ["카페명", "게시판구분", "게시글", "게시일시", "데이터추출일시", "ID", "조회수", "댓글", "좋아요", "비고"]:
        if col not in df.columns:
            df[col] = ""

    import re as _re
    URL_PATTERN = _re.compile(
        r'^https?://(cafe\.naver\.com/|naver\.me/)\S+', _re.IGNORECASE
    )

    valid_rows = []
    invalid_count = 0
    for idx, row in df.iterrows():
        url = str(row.get("업로드 링크", "")).strip()
        if url in ("", "nan"):
            continue
        if not URL_PATTERN.match(url):
            log(f"  ⚠️  유효하지 않은 URL 스킵: {url}")
            df.at[idx, "비고"] = "유효하지 않은 URL"
            df.at[idx, "데이터추출일시"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            invalid_count += 1
            continue
        valid_rows.append((idx, row))

    total = len(valid_rows)
    log(f"총 {total}개 유효 URL 처리 시작 (무효 {invalid_count}개 스킵)")
    log("-" * 50)

    # 추출중 팝업 띄우기
    popup = show_progress_popup(total)

    log("Chrome 드라이버 초기화 중...")
    popup._status[0] = "Chrome 드라이버 초기화 중..."
    try: popup.update()
    except: pass
    driver = setup_driver()
    log("Chrome 드라이버 준비 완료")
    popup._status[0] = "Chrome 드라이버 준비 완료"
    try: popup.update()
    except: pass

    try:
        for count, (idx, row) in enumerate(valid_rows, 1):
            url = str(row.get("업로드 링크", "")).strip()
            log(f"[{count}/{total}] {url}")
            popup._current[0] = count
            popup._status[0] = f"접속 중: {url[:50]}..."
            try:
                popup.update()
            except Exception:
                pass

            info = extract_post_info(driver, url)

            df.at[idx, "데이터추출일시"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            df.at[idx, "비고"]  = str(info["note"])
            # 비고가 있으면(삭제/로그인필요 등) 나머지 컬럼은 빈값 유지
            df.at[idx, "카페명"] = str(info["cafe_name"])
            if not info["note"]:
                df.at[idx, "게시판구분"] = str(info["board_name"])
                df.at[idx, "게시글"] = str(info["title"])
                df.at[idx, "게시일시"] = str(info["post_date"])
                df.at[idx, "ID"]   = str(info["author"])
                df.at[idx, "조회수"] = str(info["views"]) if info["views"] != "" else ""
                df.at[idx, "댓글"]  = str(info["comments"]) if info["comments"] != "" else ""
                df.at[idx, "좋아요"] = str(info["likes"]) if info["likes"] != "" else ""

            if info["note"]:
                log(f"  결과: {info['note']}")
            else:
                log(f"  결과: 글쓴이={info['author']} / 조회={info['views']} / 댓글={info['comments']}")
            log("-" * 50)

    finally:
        driver.quit()
        log("Chrome 드라이버 종료")
        popup._running[0] = False
        try:
            popup.destroy()
        except Exception:
            pass

    # 컬럼 순서 재정렬
    all_cols = df.columns.tolist()
    ordered = ["업로드 링크", "카페명", "게시판구분", "게시글", "게시일시", "데이터추출일시", "ID", "조회수", "댓글", "좋아요", "비고"]
    # 위 목록에 없는 나머지 컬럼은 뒤에 붙임
    rest = [c for c in all_cols if c not in ordered]
    final_cols = [c for c in ordered if c in all_cols] + rest
    df = df[final_cols]

    # 엑셀 저장 불가 특수문자 제거
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    def clean_cell(val):
        if isinstance(val, str):
            return ILLEGAL_CHARACTERS_RE.sub('', val)
        return val
    try:
        df = df.map(clean_cell)
    except AttributeError:
        df = df.applymap(clean_cell)

    # 저장 경로 탐색기
    today = datetime.now().strftime("%Y%m%d")
    base  = os.path.splitext(os.path.basename(excel_path))[0]
    default_name = f"{base}_{today}_result.xlsx"
    save_path = select_save_path(default_name)

    if save_path:
        # 엑셀 저장 + 컬럼 너비 자동 조정 + 줄바꿈
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment
        from openpyxl.utils import get_column_letter

        df.to_excel(save_path, index=False)

        wb = load_workbook(save_path)
        ws = wb.active

        # 헤더 회색 배경
        from openpyxl.styles import PatternFill, Font
        header_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col_idx, col in enumerate(ws.columns, 1):
            max_len = 0
            col_letter = get_column_letter(col_idx)
            for cell in col:
                try:
                    cell_len = max(len(str(line)) for line in str(cell.value).split("\n")) if cell.value else 0
                    max_len = max(max_len, cell_len)
                except:
                    pass
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

        wb.save(save_path)
        log(f"✅ 완료! 결과 저장: {save_path}")
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("네이버카페데이터 추출 완료", f"저장 위치:\n{save_path}")
        root.destroy()
    else:
        log("저장 취소됨")

if __name__ == "__main__":
    run()