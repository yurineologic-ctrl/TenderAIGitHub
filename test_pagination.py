#!/usr/bin/env python3
"""
Простой тест для проверки работы пагинации
"""
import time
from playwright.sync_api import sync_playwright

BASE_URL = "https://shop.lenovo.ua"
CATALOG_URL = "https://shop.lenovo.ua/category/notebooks"

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=True)
    page = browser.new_page()
    
    # Тест 1: Загрузка первой страницы
    print("Test 1: Loading page 1...")
    page.goto(CATALOG_URL, wait_until="networkidle", timeout=15_000)
    time.sleep(1)
    
    cards_1 = page.query_selector_all("article.product-card")
    print(f"  Found {len(cards_1)} cards on page 1")
    
    # Тест 2: Попытка загрузить страницу 2 прямым URL
    print("\nTest 2: Loading page 2 with ?page=2...")
    url_2 = CATALOG_URL + "?page=2"
    page.goto(url_2, wait_until="networkidle", timeout=15_000)
    time.sleep(1)
    
    cards_2 = page.query_selector_all("article.product-card")
    print(f"  Found {len(cards_2)} cards on page 2")
    
    # Тест 3: Проверка активной кнопки пагинации
    print("\nTest 3: Checking active page button...")
    active_btn = page.query_selector("button.Mui-selected.MuiPaginationItem-page")
    if active_btn:
        print(f"  Active page: {active_btn.inner_text().strip()}")
    else:
        print("  No active page button found!")
    
    # Тест 4: Попробуем другие форматы URL
    print("\nTest 4: Trying other URL formats...")
    for pattern in [
        CATALOG_URL + "&page=3",
        CATALOG_URL + "/page/3",
        CATALOG_URL + "/3",
    ]:
        print(f"  Trying: {pattern}")
        try:
            page.goto(pattern, wait_until="networkidle", timeout=10_000)
            time.sleep(0.5)
            cards = page.query_selector_all("article.product-card")
            print(f"    → Found {len(cards)} cards")
            if len(cards) > 0:
                break
        except Exception as e:
            print(f"    → Error: {e}")
    
    browser.close()

print("\n✓ Test complete!")
