# -*- coding: utf-8 -*-
"""dsh web E2E：发起对话，让模型调用 excel_describe 分析 small.xlsx"""
import os, sys, json, time
from playwright.sync_api import sync_playwright

URL = "http://127.0.0.1:3080"
# 仓库根目录 = test/ 的上一级；用相对路径定位 fixture 与截图，换机器可跑
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(BASE, "test", "fixtures", "small.xlsx")
SHOT1 = os.path.join(BASE, "test", "e2e-01-home.png")
SHOT2 = os.path.join(BASE, "test", "e2e-02-sent.png")
SHOT3 = os.path.join(BASE, "test", "e2e-03-result.png")
PROMPT = ("请使用 excel_describe 工具分析这个 Excel 文件：" + XLSX +
          "。请调用工具并告诉我：sheet 列表、总行数、总列数、每列的空值率，以及数值列的 min/max/mean。")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1600, "height": 1000})
    page.goto(URL, wait_until="domcontentloaded", timeout=20000)
    page.wait_for_timeout(4000)
    page.screenshot(path=SHOT1)
    print("STEP1_OK title=", page.title())

    # 找输入框：优先 TEXTAREA（对话输入），跳过搜索框
    sel = page.evaluate("""() => {
        const q = Array.from(document.querySelectorAll('textarea, [contenteditable="true"]'));
        if (q.length === 0) return null;
        const el = q[0];
        el.focus();
        return {tag: el.tagName, ce: el.getAttribute('contenteditable'), ph: el.placeholder || ''};
    }""")
    if sel is None:
        print("NO_INPUT_FOUND")
        body = page.evaluate("() => document.body.innerText.slice(0, 1500)")
        print("BODY:", body)
        browser.close()
        sys.exit(4)
    print("INPUT:", json.dumps(sel, ensure_ascii=False))

    # 输入提示词（textarea 用 value 赋值）
    typed = page.evaluate("""(text) => {
        const q = document.querySelectorAll('textarea, [contenteditable="true"]');
        const el = q[0];
        if (el.tagName === 'TEXTAREA') {
            const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
            setter.call(el, text);
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
        } else {
            el.textContent = text;
            el.dispatchEvent(new InputEvent('input', {bubbles: true, data: text}));
        }
        return el.value !== undefined ? el.value.length : text.length;
    }""", PROMPT)
    print("TYPED_CHARS:", typed)
    page.wait_for_timeout(800)
    page.screenshot(path=SHOT2)

    # 点击发送按钮（aria-label=发送消息）
    try:
        clicked = page.evaluate("""() => {
            const b = Array.from(document.querySelectorAll('button')).find(x => (x.getAttribute('aria-label')||'').includes('发送'));
            if (b) { b.click(); return true; }
            return false;
        }""")
        print("SEND_BUTTON_CLICKED:", clicked)
    except Exception as e:
        print("SEND_FAIL:", str(e)[:100])
    if not clicked:
        page.keyboard.press("Enter")

    # 等待模型响应（最多 100s）
    print("WAITING_FOR_MODEL...")
    for i in range(5):
        page.wait_for_timeout(4000)
        body = page.evaluate("() => document.body.innerText")
        low = body.lower()
        # 检查是否出现错误/工具调用迹象
        if any(k in low for k in ["excel_describe", "error", "api key", "401", "unauthorized", "invalid", "failed", "失败", "错误"]):
            print(f"T{i*5}s SIGNAL_DETECTED")
            break
    page.screenshot(path=SHOT3)
    body = page.evaluate("() => document.body.innerText")
    print("===FINAL_BODY(3000)===")
    print(body[-3000:] if len(body) > 3000 else body)
    browser.close()
    print("DONE")
