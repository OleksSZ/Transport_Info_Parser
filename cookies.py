
import time
import json
import undetected_chromedriver as uc

def parse_cookies(url="https://brokersnapshot.com/"):
    """Открывает браузер без обнаружения, даёт 60 секунд на логин, сохраняет cookies"""
    print("Запускаю undetected_chromedriver...")

    # Настройки для обхода обнаружения
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = uc.Chrome(options=options, use_subprocess=True)
    driver.get(url)

    print("Браузер открыт. У вас 60 секунд на регистрацию/вход.")
    time.sleep(60)

    # Сохраняем cookies
    cookies = driver.get_cookies()
    with open("cookies.json", "w", encoding="utf-8") as f:
        json.dump(cookies, f, indent=4, ensure_ascii=False)

    print("Cookies сохранены в cookies.json")
    driver.quit()

if __name__ == "__main__":
    parse_cookies()