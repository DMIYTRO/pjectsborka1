import datetime
import locale
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import psutil
for proc in psutil.process_iter(['name']):
    if proc.info['name'] == 'chromedriver.exe':
        proc.kill()

login_admin = "admin"
password = ""
login_user = "1kuperster"

open_window_brauser = False
# download_path = "rC:/Users/Dmitry/Downloads/сборки_xls"
my_time = "10.09, Чт"
local_day = locale.setlocale(locale.LC_TIME, "ru_RU")
now = datetime.datetime.today().strftime("%d.%m, %a")
same_date = now.title()
print(same_date)
# driver = webdriver.Chrome(options=chrome_options)
test_time = "01:55"


def configure_browser(download_path):
    options = Options()
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-extensions")
    options.add_argument("--proxy-server='direct://'")
    options.add_argument("--proxy-bypass-list=*")
    options.add_argument("--start-maximized")
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--no-sandbox')
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko)"
                         " Chrome/92.0.4515.107 Safari/537.36")
    # options.add_argument("download.default_directory: download_path")
    options.headless = open_window_brauser  # скрыть окно браузера
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko)"
                         " Chrome/92.0.4515.107 Safari/537.36")
    options.add_experimental_option("prefs", {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    driver = webdriver.Chrome(options=options, executable_path=r'C:\Users\Dmitry\PycharmProjects\pythonProject\chromedriver.exe')
    # Set the download directory preferenc
    return driver, options


download_path = r"C:\Users\Dmitry\Downloads\сборки_xls"


def new():
    driver, options = configure_browser(download_path)
    url = "https://admin:R5PD0+}Qzqr@@www.sborka.ua/adm/sborki.php?type=digital"
    driver.get(url)
    assert "Сборка. Панель работника" in driver.title
    driver.find_element_by_name('pass_worker').send_keys(login_user)
    time.sleep(1)
    driver.find_element_by_name('button').click()
    print('1')
    time.sleep(1)
    driver.find_element_by_xpath("//a[@href='sborki.php']").click()
    time.sleep(1)
    driver.find_element_by_xpath("//a[@href='?type=ofset']").click()
    time.sleep(1)

    # Найдите все элементы td
    td_elements = driver.find_elements_by_tag_name("td")
    for td_element in td_elements:
        if td_element.text in ["Логос", "Ротекс", "Аванпост"]:
            try:
                img_element = td_element.find_element_by_xpath("./following-sibling::td/img[@alt='уже']")
                tr_elements = driver.find_elements_by_xpath(
                    "//tr[./td[text()='Логос' or text()='Ротекс' or text()='Аванпост']]")
                num_elements = len(tr_elements)  # Подсчет количества элементов
                for i in range(num_elements):
                    report_img_elements = tr_elements[i].find_elements_by_xpath(
                        ".//a[@title='экспорт в эксель']//img[@src='../images/report-excel.png']")
                    for report_img_element in report_img_elements:
                        # Далее можно выполнять операции с найденными элементами
                        print(2)  # ...
                        time.sleep(0)
                        report_img_element.click()
                        time.sleep(0)
                break  # Выход из цикла после обработки всех элементов
            except NoSuchElementException:
                print("Элемент не отмечен ")
                continue  # Продолжить цикл for
        else:
            continue  # Перейти к следующему элементу td
    time.sleep(3)
    driver.quit()


if __name__ == '__main__':
    new()



