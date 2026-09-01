import os
import time
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ===== 유틸 =====
def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)

# ===== 상수 =====
RANK_URL = "https://m.oliveyoung.co.kr/m/mtn?menu=ranking&tab=brands&timeSaleDayFilter=tomorrow"
OUT_XLSX = "oliveyoung_brand_rankings_wide.xlsx"
PAGE_LOAD_TIMEOUT = 25

# 스크롤/수집 파라미터(느리게!)
SCROLL_STEP_RATIO = 0.45     # 한 번에 내릴 비율(작게)
SLEEP_AFTER_PARSE = 0.28     # 파싱 직후 대기
SLEEP_AFTER_SCROLL = 0.28    # 스크롤 직후 대기
NO_NEW_COOLDOWN = 0.6        # 새 항목이 안 늘어났을 때 추가 대기
NO_NEW_BUMP_SEC = 3.0        # 새 항목 증가가 멈춘 것으로 보는 기준(sec)
MAX_DOWN_TRIES = 1600
UP_PASS_STEPS = 120          # 100위 도달 후 위로 훑는 스텝 수

# (선택) 크롬 사용자 프로필 경로: 사용 중이면 자동 폴백됨
CHROME_USER_DATA_DIR = r"C:\Users\user\AppData\Local\Google\Chrome\User Data"  # 사용 안할 땐 None
CHROME_PROFILE_DIR = "Default"

# ===== 스크롤 도우미 =====
def robust_scroll_down(driver, ratio=0.9):
    """
    - 1순위: 가상리스트/메인 스크롤러 scrollTop 직접 증가
    - 2순위: Selenium ActionChains.scroll_by_amount
    - 3순위: CDP MouseWheel
    """
    js = r"""
    (function(ratio){
      var step = (ratio>1) ? Math.floor(ratio) : Math.floor(window.innerHeight*ratio);
      var out = {moved:false, used:null, before:0, after:0, max:0, step:step};
      var winEl = document.scrollingElement || document.documentElement || document.body;

      var cand = [
        winEl,
        document.querySelector('[data-virtuoso-scroller="true"]'),
        document.querySelector('.swiper-slide-active'),
        document.querySelector('#main-inner-swiper-ranking')
      ].filter(Boolean);

      for (var i=0; i<cand.length; i++){
        var el = cand[i];
        var before = el.scrollTop || 0;
        var max = (el.scrollHeight - el.clientHeight);
        el.scrollTop = Math.min(before + step, el.scrollHeight);
        var after = el.scrollTop || 0;
        out.before = before; out.after = after; out.max = max;
        if (after > before || (max > 0 && after >= max)) {
          out.moved = true;
          out.used = (el===winEl ? 'document.scrollingElement' : (el.className || el.id || 'node'));
          break;
        }
      }

      if (!out.moved) {
        window.scrollBy(0, step);
        var b2 = winEl.scrollTop || 0;
        if (b2 > out.before) {
          out.moved = true; out.used = 'window.scrollBy'; out.after = b2;
        }
      }
      return out;
    })(arguments[0]);
    """
    try:
        info = driver.execute_script(js, ratio)
        if info and info.get("moved"):
            return True
    except Exception:
        pass

    try:
        h = driver.get_window_size()["height"]
        dy = int(h * (ratio if ratio <= 1 else 1))
        ActionChains(driver).scroll_by_amount(0, dy).perform()
        return True
    except Exception:
        pass

    try:
        driver.execute_cdp_cmd("Input.dispatchMouseWheelEvent", {
            "x": 200, "y": 600, "deltaX": 0, "deltaY": 400
        })
        return True
    except Exception:
        pass

    return False

def robust_scroll_up(driver, ratio=0.9):
    # 위로 올릴 때도 같은 후보들을 역방향
    js = r"""
    (function(ratio){
      var step = (ratio>1) ? Math.floor(ratio) : Math.floor(window.innerHeight*ratio);
      var out = {moved:false, used:null, before:0, after:0, step:step};
      var winEl = document.scrollingElement || document.documentElement || document.body;

      var cand = [
        document.querySelector('[data-virtuoso-scroller="true"]'),
        winEl
      ].filter(Boolean);

      for (var i=0; i<cand.length; i++){
        var el = cand[i];
        var before = el.scrollTop || 0;
        el.scrollTop = Math.max(before - step, 0);
        var after = el.scrollTop || 0;
        out.before = before; out.after = after;
        if (after < before || after === 0) {
          out.moved = true;
          out.used = (el===winEl ? 'document.scrollingElement' : (el.className || el.id || 'node'));
          break;
        }
      }
      if (!out.moved) {
        window.scrollBy(0, -step);
        var b2 = winEl.scrollTop || 0;
        if (b2 < out.before || b2 == 0) {
          out.moved = true; out.used = 'window.scrollBy'; out.after = b2;
        }
      }
      return out;
    })(arguments[0]);
    """
    try:
        info = driver.execute_script(js, ratio)
        if info and info.get("moved"):
            return True
    except Exception:
        pass
    try:
        h = driver.get_window_size()["height"]
        dy = int(h * (ratio if ratio <= 1 else 1))
        ActionChains(driver).scroll_by_amount(0, -dy).perform()
        return True
    except Exception:
        pass
    try:
        driver.execute_cdp_cmd("Input.dispatchMouseWheelEvent", {
            "x": 200, "y": 600, "deltaX": 0, "deltaY": -400
        })
        return True
    except Exception:
        pass
    return False

# ===== 드라이버: 모바일 에뮬레이션 =====
def build_driver(user_data_dir=None, profile_dir=None):
    options = Options()
    # options.add_argument("--headless=new")  # 필요시
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    mobile_emulation = {
        "deviceMetrics": {"width": 390, "height": 844, "pixelRatio": 3},
        "userAgent": (
            "Mozilla/5.0 (Linux; Android 12; SM-G996N) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/139.0.0.0 Mobile Safari/537.36"
        )
    }
    options.add_experimental_option("mobileEmulation", mobile_emulation)

    if user_data_dir:
        options.add_argument(f"--user-data-dir={user_data_dir}")
    if profile_dir:
        options.add_argument(f"--profile-directory={profile_dir}")

    service = ChromeService()
    driver = webdriver.Chrome(options=options, service=service)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
    })
    return driver

# ===== 진입 & 안정 대기 =====
def goto_ranking_page(driver):
    log("브랜드 랭킹 페이지 접속 중…")
    driver.get(RANK_URL)

    WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    # 브랜드 탭이 아닐 경우 클릭
    try:
        brands_tab = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a[data-tab="brands"]'))
        )
        cls = brands_tab.get_attribute("class") or ""
        if "TabItem_active" not in cls:
            driver.execute_script("arguments[0].click()", brands_tab)
            time.sleep(0.8)
    except Exception:
        pass

    # 리스트 루트 등장 대기
    WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
        lambda d: (
            d.find_elements(By.CSS_SELECTOR, '[class^="BestBrandsList_brand__"]')
            or d.find_elements(By.CSS_SELECTOR, '[class^="BestBrandsList_rank__"] strong')
            or d.find_elements(By.CSS_SELECTOR, '[data-virtuoso-scroller="true"]')
        )
    )

    # 스크롤러 포커스
    try:
        scroller = driver.find_element(By.CSS_SELECTOR, '[data-virtuoso-scroller="true"]')
        driver.execute_script("arguments[0].scrollTop = 0;", scroller)
        driver.execute_script("try{arguments[0].focus()}catch(e){}", scroller)
    except Exception:
        pass

    log("리스트 감지 완료.")

# ===== 가시 영역 파싱 =====
def parse_visible_pairs(driver):
    """현재 DOM에 그려진 카드에서 (rank, name) 수집"""
    pairs = []
    cards = driver.find_elements(By.CSS_SELECTOR, '[class^="BestBrandsList_brand__"]')
    for card in cards:
        try:
            rank_el = card.find_element(By.CSS_SELECTOR, '[class^="BestBrandsList_rank__"] strong')
            name_el = card.find_element(By.CSS_SELECTOR, '[class^="BestBrandsList_header__"] p')
            rank = rank_el.text.strip()
            name = name_el.text.strip()
            if rank.isdigit():
                pairs.append((int(rank), name))
        except Exception:
            continue
    return pairs

# ===== 수집 & 저장 =====
def collect_and_save(driver):
    log("느리게 스크롤하며 수집 시작…")

    dedup = {}               # {rank: name}
    seen_max_rank = 0
    last_new_time = time.time()

    # 아래로 천천히 훑으면서 계속 수집
    for i in range(MAX_DOWN_TRIES):
        # 1) 현재 보이는 항목 파싱
        pairs = parse_visible_pairs(driver)
        added = 0
        for r, n in pairs:
            if 1 <= r <= 100 and r not in dedup:
                dedup[r] = n
                added += 1
                if r > seen_max_rank:
                    seen_max_rank = r
        if added:
            last_new_time = time.time()

        # 진행 로그(가끔)
        if i % 30 == 0:
            log(f"진행: {len(dedup)}개 수집, 현재 최대 랭크 {seen_max_rank}")

        # 100위 감지 시, 아래서 위로 보정 패스 준비
        if seen_max_rank >= 100 and len(dedup) >= 60:  # 어느정도 채웠으면
            log("100위 감지! 아래→위 보정 패스 진행…")
            time.sleep(0.8)
            break

        # 2) 잠깐 숨고르기
        time.sleep(SLEEP_AFTER_PARSE)

        # 3) 천천히 스크롤
        robust_scroll_down(driver, SCROLL_STEP_RATIO)
        time.sleep(SLEEP_AFTER_SCROLL)

        # 4) 새 항목이 일정 시간 안 늘어나면 더 기다렸다가 다음 스크롤
        if time.time() - last_new_time > NO_NEW_BUMP_SEC:
            time.sleep(NO_NEW_COOLDOWN)
            last_new_time = time.time()

    # 보정 패스: 아래에서 위로 천천히 올리며 누락 수집
    if seen_max_rank >= 100:
        for _ in range(UP_PASS_STEPS):
            pairs = parse_visible_pairs(driver)
            added = 0
            for r, n in pairs:
                if 1 <= r <= 100 and r not in dedup:
                    dedup[r] = n
                    added += 1
            if added:
                last_new_time = time.time()
            time.sleep(SLEEP_AFTER_PARSE)
            robust_scroll_up(driver, SCROLL_STEP_RATIO * 0.9)
            time.sleep(SLEEP_AFTER_SCROLL)
            if len(dedup) >= 100:
                break

    # 최종 정리
    final = [(r, dedup[r]) for r in sorted(dedup.keys())]
    log(f"수집 개수: {len(final)}")

    # ==== 엑셀 저장 (가로형 유지) ====
    cols = ["크롤링날짜", "요일"] + [f"{i}위" for i in range(1, 101)]
    today = datetime.now()
    weekday_kor = "월화수목금토일"[today.weekday()]
    row = {"크롤링날짜": today.strftime("%Y-%m-%d"), "요일": weekday_kor}
    for r, n in final:
        row[f"{r}위"] = n

    df_new = pd.DataFrame([row], columns=cols)

    if os.path.exists(OUT_XLSX):
        try:
            df_old = pd.read_excel(OUT_XLSX)
            for c in cols:
                if c not in df_old.columns:
                    df_old[c] = None
            df_all = pd.concat([df_old[cols], df_new], ignore_index=True)
        except Exception:
            df_all = df_new
        df_all.to_excel(OUT_XLSX, index=False)
    else:
        df_new.to_excel(OUT_XLSX, index=False)

    log(f"엑셀 저장 완료 → {os.path.abspath(OUT_XLSX)}")

# ===== 메인 =====
def main():
    driver = None
    try:
        if CHROME_USER_DATA_DIR:
            log("크롬: 사용자 프로필로 시도 중...")
            try:
                driver = build_driver(CHROME_USER_DATA_DIR, CHROME_PROFILE_DIR)
                log("기존 브라우저 세션에서 여는 중입니다.")
            except Exception as e:
                log(f"[경고] 사용자 프로필 실패 → 임시 프로필로 전환: {getattr(e, 'msg', str(e))}")
                driver = build_driver()
        else:
            log("크롬: 임시 프로필 사용 (user-data-dir 미지정)")
            driver = build_driver()

        goto_ranking_page(driver)
        collect_and_save(driver)
        log("🎉 완료")
    finally:
        try:
            if driver:
                driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()
