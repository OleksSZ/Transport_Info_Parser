# Dot_pars.py — с гарантированным закрытием браузера

import openpyxl
import time
import json
import os
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
import re
import traceback
import psutil

def extract_company_dots(url, excel_path="company_dots.xlsx"):
    driver = None
    print(f"Запуск Dot_pars.py для URL: {url}")

    try:
        print("Создаю опции...")
        options = uc.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")

        print("Запускаю браузер...")
        driver = uc.Chrome(
            options=options,
            use_subprocess=True
        )
        print("Браузер запущен")

        # Cookies
        if os.path.exists("cookies.json"):
            print("Загружаю cookies...")
            driver.get("https://brokersnapshot.com/")
            with open("cookies.json", "r", encoding="utf-8") as f:
                cookies = json.load(f)
                for cookie in cookies:
                    cookie.pop('expiry', None)
                    cookie.pop('sameSite', None)
                    try:
                        driver.add_cookie(cookie)
                    except:
                        pass
            driver.refresh()
            time.sleep(3)
            print("Cookies применены")

        print(f"Перехожу на страницу: {url}")
        driver.get(url)
        driver.implicitly_wait(20)
        time.sleep(12)

        print("Ищу строки таблицы...")
        rows = driver.find_elements(By.XPATH, "//tr[td[@data-label]]")
        if not rows:
            print("Строк не найдено! Проверяю альтернативный XPath...")
            rows = driver.find_elements(By.TAG_NAME, "tr")
            print(f"Найдено tr: {len(rows)}")

        results = []
        for row in rows:
            try:
                # DOT
                dot_td = row.find_element(By.XPATH, ".//td[@data-label='DOT']")
                dot_text = dot_td.text.strip()
                dot_clean = re.sub(r'[^0-9]', '', dot_text)

                # Link
                try:
                    link_elem = row.find_element(By.XPATH, ".//a[contains(@href, '/Company')]")
                    link = link_elem.get_attribute("href")
                    if not link.startswith("http"):
                        link = "https://brokersnapshot.com" + link
                except:
                    link = "-"

                if dot_clean.isdigit():
                    results.append({"DOT": dot_clean, "Link": link})
            except:
                continue

        print(f"Найдено компаний: {len(results)}")

        # Сохранение
        try:
            wb = openpyxl.load_workbook(excel_path)
            sheet = wb.active
        except FileNotFoundError:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.append(["DOT", "Link"])

        existing = {str(cell.value) for cell in sheet['A'] if cell.value}
        added_count = 0
        for item in results:
            if item["DOT"] not in existing:
                sheet.append([item["DOT"], item["Link"]])
                existing.add(item["DOT"])
                added_count += 1

        wb.save(excel_path)
        print(f"Сохранено в {excel_path}")

        return f"✅ Найдено {len(results)} компаний\nДобавлено новых: {added_count}"

    except Exception as e:
        print("Критическая ошибка в Dot_pars:")
        traceback.print_exc()
        return f"Ошибка: {str(e)}"

    finally:
        print("Закрываю браузер...")
        if driver is not None:
            try:
                driver.quit()
                print("driver.quit() выполнен")
            except Exception as quit_err:
                print(f"driver.quit() не сработал: {quit_err}")
                # Принудительно убиваем все процессы chromedriver и chrome
                try:
                    for proc in psutil.process_iter(['name']):
                        if proc.name() in ("chromedriver.exe", "chrome.exe"):
                            proc.kill()
                            print(f"Убит процесс: {proc.name()} (PID {proc.pid})")
                except Exception as kill_err:
                    print(f"Ошибка убийства процессов: {kill_err}")

        print("Закрытие завершено")