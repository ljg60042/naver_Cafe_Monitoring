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
import sys
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


def get_app_dir():
    """실행 파일(exe) 기준 또는 스크립트 기준 폴더 경로 반환"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


NICKNAME_FILE = os.path.join(get_app_dir(), "nicknames.txt")


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


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
    result = {"author": "", "views": "", "comments": "", "likes": "", "note": "", "post_date": "", "cafe_name": "", "board_name": "", "title": "", "comments_list": []}

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

        # ── 댓글 목록 (작성자 + 내용) ────────────────────────
        log("  댓글 목록 추출 중...")
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "ul.comment_list li.CommentItem"))
            )
        except TimeoutException:
            pass  # 댓글이 없는 게시글일 수 있음

        try:
            comment_items = driver.find_elements(By.CSS_SELECTOR, "li.CommentItem")
            comments_list = []
            for item in comment_items:
                try:
                    nick = item.find_element(By.CSS_SELECTOR, "a.comment_nickname").text.strip()
                except NoSuchElementException:
                    nick = ""
                try:
                    comment_text = item.find_element(By.CSS_SELECTOR, ".comment_text_box").text.strip()
                except NoSuchElementException:
                    comment_text = ""
                if nick:
                    comments_list.append({"author": nick, "text": comment_text})
            result["comments_list"] = comments_list
            log(f"  댓글 {len(comments_list)}개 추출 (작성자: {[c['author'] for c in comments_list]})")
        except Exception:
            log("  ⚠️  댓글 목록 추출 실패")

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
            m = re.search(r"Alert Text: (.+?)(\n|Message:)", err)
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


class NaverCafeMonitorApp:
    """엑셀 파일 선택 → 모니터링 시작 버튼으로 동작하는 메인 UI"""

    def __init__(self):
        self.excel_path = None

        self.root = tk.Tk()
        self.root.title("네이버 카페 모니터링")
        self.root.geometry("420x460")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 420) // 2
        y = (self.root.winfo_screenheight() - 460) // 2
        self.root.geometry(f"420x460+{x}+{y}")

        tk.Label(self.root, text="네이버 카페 모니터링", font=("맑은 고딕", 14, "bold")).pack(pady=(20, 15))

        tk.Label(self.root, text="추적할 네이버 아이디(닉네임) 목록").pack()

        id_frame = tk.Frame(self.root)
        id_frame.pack(pady=(5, 5))
        self.id_var = tk.StringVar()
        self.id_entry = tk.Entry(id_frame, textvariable=self.id_var, width=22)
        self.id_entry.pack(side="left", padx=(0, 6))
        self.id_entry.bind("<Return>", lambda e: self.add_nickname())
        self.add_btn = tk.Button(id_frame, text="추가", width=6, command=self.add_nickname)
        self.add_btn.pack(side="left")

        list_frame = tk.Frame(self.root)
        list_frame.pack(pady=(0, 5))
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        self.nickname_listbox = tk.Listbox(list_frame, height=6, width=34, yscrollcommand=scrollbar.set)
        self.nickname_listbox.pack(side="left")
        scrollbar.config(command=self.nickname_listbox.yview)

        self.remove_btn = tk.Button(self.root, text="선택 삭제", width=12, command=self.remove_nickname)
        self.remove_btn.pack(pady=(0, 12))

        self.file_var = tk.StringVar(value="선택된 엑셀 파일이 없습니다")
        tk.Label(self.root, textvariable=self.file_var, fg="#555555", wraplength=380, justify="center").pack(pady=(0, 10))

        self.select_file_btn = tk.Button(self.root, text="엑셀 파일 선택", width=22, command=self.select_file)
        self.select_file_btn.pack(pady=5)

        self.start_btn = tk.Button(
            self.root, text="모니터링 시작", width=22, state="disabled",
            bg="#4CAF50", fg="white", command=self.start_monitoring
        )
        self.start_btn.pack(pady=5)

        self.status_var = tk.StringVar(value="")
        tk.Label(self.root, textvariable=self.status_var, fg="#333333", font=("맑은 고딕", 10)).pack(pady=(15, 0))

        self.progress_var = tk.StringVar(value="")
        tk.Label(self.root, textvariable=self.progress_var, fg="#777777", font=("맑은 고딕", 10)).pack()

        self.load_nicknames()

    def select_file(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title="엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]
        )
        if path:
            self.excel_path = path
            self.file_var.set(f"선택됨: {os.path.basename(path)}")
            self.update_start_button_state()

    def add_nickname(self):
        nickname = self.id_var.get().strip()
        if not nickname:
            return
        existing = self.nickname_listbox.get(0, tk.END)
        if nickname in existing:
            messagebox.showinfo("알림", "이미 추가된 아이디입니다.", parent=self.root)
            return
        self.nickname_listbox.insert(tk.END, nickname)
        self.id_var.set("")
        self.update_start_button_state()
        self.save_nicknames()

    def remove_nickname(self):
        selection = self.nickname_listbox.curselection()
        if not selection:
            messagebox.showwarning("알림", "삭제할 아이디를 목록에서 선택해주세요.", parent=self.root)
            return
        for i in reversed(selection):
            self.nickname_listbox.delete(i)
        self.update_start_button_state()
        self.save_nicknames()

    def get_nickname_list(self):
        return list(self.nickname_listbox.get(0, tk.END))

    def load_nicknames(self):
        """nicknames.txt에서 이전에 저장한 닉네임 목록을 불러옴"""
        if not os.path.exists(NICKNAME_FILE):
            return
        try:
            with open(NICKNAME_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    nick = line.strip()
                    if nick:
                        self.nickname_listbox.insert(tk.END, nick)
            self.update_start_button_state()
        except Exception as e:
            log(f"⚠️  닉네임 파일 로드 실패: {e}")

    def save_nicknames(self):
        """현재 닉네임 목록을 nicknames.txt에 저장"""
        try:
            with open(NICKNAME_FILE, "w", encoding="utf-8") as f:
                for nick in self.nickname_listbox.get(0, tk.END):
                    f.write(nick + "\n")
        except Exception as e:
            log(f"⚠️  닉네임 파일 저장 실패: {e}")
            messagebox.showwarning("알림", f"닉네임 저장에 실패했습니다:\n{e}", parent=self.root)

    def update_start_button_state(self):
        if self.excel_path and self.nickname_listbox.size() > 0:
            self.start_btn.config(state="normal")
        else:
            self.start_btn.config(state="disabled")

    def set_controls_enabled(self, enabled):
        """모니터링 진행 중에는 다른 입력/버튼을 잠가서 조작 못 하게 함"""
        state = "normal" if enabled else "disabled"
        self.id_entry.config(state=state)
        self.add_btn.config(state=state)
        self.remove_btn.config(state=state)
        self.nickname_listbox.config(state=state)
        self.select_file_btn.config(state=state)

    def start_monitoring(self):
        if not self.excel_path:
            messagebox.showwarning("알림", "먼저 엑셀 파일을 선택해주세요.", parent=self.root)
            return
        if self.nickname_listbox.size() == 0:
            messagebox.showwarning("알림", "네이버 아이디(닉네임)를 하나 이상 추가해주세요.", parent=self.root)
            return
        self.start_btn.config(state="disabled")
        self.set_controls_enabled(False)
        try:
            self.run_monitoring()
        except Exception as e:
            log(f"❌ 모니터링 중 오류 발생: {e}")
            messagebox.showerror("오류", f"모니터링 중 오류가 발생했습니다:\n{e}", parent=self.root)
        finally:
            self.set_controls_enabled(True)
            self.update_start_button_state()

    def run_monitoring(self):
        excel_path = self.excel_path
        target_ids = self.get_nickname_list()
        folder = os.path.dirname(excel_path)
        base = os.path.splitext(os.path.basename(excel_path))[0]

        log("=" * 50)
        log("네이버 카페 모니터링 시작")
        log("=" * 50)
        log(f"파일: {excel_path}")
        log(f"대상 아이디 목록: {', '.join(target_ids)}")

        df = pd.read_excel(excel_path, dtype=str)
        df.columns = df.columns.str.strip()

        for col in ["카페명", "게시판구분", "게시글", "게시일시", "데이터추출일시", "ID", "조회수", "댓글", "좋아요", "비고"]:
            if col not in df.columns:
                df[col] = ""

        URL_PATTERN = re.compile(
            r'^https?://(cafe\.naver\.com/|naver\.me/)\S+', re.IGNORECASE
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

        self.status_var.set("Chrome 드라이버 초기화 중...")
        self.progress_var.set(f"[0 / {total}]")
        self.root.update()

        driver = setup_driver()
        log("Chrome 드라이버 준비 완료")
        self.status_var.set("Chrome 드라이버 준비 완료")
        self.root.update()

        comments_by_idx = {}  # idx -> 등록된 아이디가 작성한 댓글 목록 (행 분리용)

        try:
            for count, (idx, row) in enumerate(valid_rows, 1):
                url = str(row.get("업로드 링크", "")).strip()
                log(f"[{count}/{total}] {url}")
                self.progress_var.set(f"[{count} / {total}]")
                self.status_var.set(f"접속 중: {url[:40]}...")
                self.root.update()

                info = extract_post_info(driver, url)

                df.at[idx, "데이터추출일시"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                df.at[idx, "비고"] = str(info["note"])
                # 비고가 있으면(삭제/로그인필요 등) 나머지 컬럼은 빈값 유지
                df.at[idx, "카페명"] = str(info["cafe_name"])
                if not info["note"]:
                    df.at[idx, "게시판구분"] = str(info["board_name"])
                    df.at[idx, "게시글"] = str(info["title"])
                    df.at[idx, "게시일시"] = str(info["post_date"])
                    df.at[idx, "ID"] = str(info["author"])
                    df.at[idx, "조회수"] = str(info["views"]) if info["views"] != "" else ""
                    df.at[idx, "댓글"] = str(info["comments"]) if info["comments"] != "" else ""
                    df.at[idx, "좋아요"] = str(info["likes"]) if info["likes"] != "" else ""

                    matched_comments = [
                        c for c in info["comments_list"]
                        if c["author"] in target_ids and c["text"]
                    ]
                    if matched_comments:
                        comments_by_idx[idx] = matched_comments

                if info["note"]:
                    log(f"  결과: {info['note']}")
                else:
                    log(f"  결과: 글쓴이={info['author']} / 조회={info['views']} / 댓글={info['comments']}")
                log("-" * 50)
        finally:
            driver.quit()
            log("Chrome 드라이버 종료")

        # 댓글이 여러 개 일치하면 댓글 하나당 한 행으로 분리
        output_records = []
        for idx, row in df.iterrows():
            row_dict = row.to_dict()
            matched = comments_by_idx.get(idx, [])
            if matched:
                for i, c in enumerate(matched, 1):
                    if i == 1:
                        record = dict(row_dict)
                    else:
                        # 두 번째 행부터는 게시글 정보는 비우고 댓글 정보만 표시
                        record = {col: "" for col in row_dict}
                    record["댓글순번"] = i
                    record["댓글닉네임"] = c["author"]
                    record["댓글내용"] = c["text"]
                    output_records.append(record)
            else:
                record = dict(row_dict)
                record["댓글순번"] = ""
                record["댓글닉네임"] = ""
                record["댓글내용"] = ""
                output_records.append(record)
        df = pd.DataFrame(output_records)

        # 컬럼 순서 재정렬
        all_cols = df.columns.tolist()
        ordered = ["업로드 링크", "카페명", "게시판구분", "게시글", "게시일시", "데이터추출일시", "ID", "조회수", "댓글", "좋아요", "댓글순번", "댓글닉네임", "댓글내용", "비고"]
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
        self.status_var.set("결과 저장 위치를 선택해주세요")
        self.root.update()

        today = datetime.now().strftime("%Y%m%d")
        default_name = f"{base}_{today}_result.xlsx"
        save_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="결과 파일 저장",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")]
        )

        if save_path:
            # 엑셀 저장 + 컬럼 너비 자동 조정 + 줄바꿈
            from openpyxl import load_workbook
            from openpyxl.styles import Alignment, PatternFill, Font
            from openpyxl.utils import get_column_letter

            df.to_excel(save_path, index=False)

            wb = load_workbook(save_path)
            ws = wb.active

            # 헤더 회색 배경
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
                    except Exception:
                        pass
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

            wb.save(save_path)
            log(f"✅ 완료! 결과 저장: {save_path}")
            self.status_var.set("완료!")
            self.progress_var.set(f"저장 위치: {os.path.basename(save_path)}")
            messagebox.showinfo("네이버카페데이터 추출 완료", f"저장 위치:\n{save_path}", parent=self.root)
        else:
            log("저장 취소됨")
            self.status_var.set("저장 취소됨")

        # 다음 모니터링을 위해 초기화 (아이디는 유지해서 다른 파일 이어서 확인 가능)
        self.excel_path = None
        self.file_var.set("선택된 엑셀 파일이 없습니다")
        self.set_controls_enabled(True)
        self.update_start_button_state()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = NaverCafeMonitorApp()
    app.run()
