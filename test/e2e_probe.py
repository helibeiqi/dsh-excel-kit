# -*- coding: utf-8 -*-
"""dsh web UI 探测：打开页面、截图、提取文本结构"""
import sys, json
from playwright.sync_api import sync_playwright

URL = "http://127.0.0.1:3080"
SHOT = r"C:\Users\ethan.he\WorkBuddy\2026-08-18-09-23-32\dsh-excel-kit\test\e2e-01-home.png"

with sync_playwright() as p:
    try:
        browser = p.chromium.launch(headless=True)
    except Exception as e:
        print("LAUNCH_FAIL:", str(e)[:300])
        sys.exit(3)
    page = browser.new_page(viewport={"width": 1600, "height": 1000})
    page.goto(URL, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(3000)
    page.screenshot(path=SHOT, full_page=False)
    title = page.title()
    print("TITLE:", title)
    print("URL:", page.url)
    # 提取主要可见文本（前 2000 字符）
    text = page.evaluate("() => document.body ? document.body.innerText.slice(0, 2000) : ''")
    print("BODY_TEXT_START:")
    print(text)
    # 找输入框
    inputs = page.evaluate("""() => Array.from(document.querySelectorAll('textarea, input[type=text], [contenteditable=true]')).slice(0,10).map((el,i)=>({i, tag: el.tagName, ph: el.placeholder||'', cls: (el.className||'').toString().slice(0,60)}))""")
    print("INPUTS:", json.dumps(inputs, ensure_ascii=False))
    browser.close()
    print("SHOT_SAVED:", SHOT)
