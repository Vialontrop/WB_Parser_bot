from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import logging
import time
import re
import pandas as pd
from selenium.common.exceptions import WebDriverException
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
# Настройка логирования
logging.basicConfig(level=logging.INFO)

def smooth_scroll_to_bottom(driver):
    """
    Плавно прокручивает страницу до конца, чтобы загрузить все элементы.
    :param driver: WebDriver Selenium.
    """
    scroll_pause_time = 0.5  # Уменьшаем паузу между прокрутками
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        # Прокручиваем страницу на небольшую величину (например, 900 пикселей)
        driver.execute_script("window.scrollBy(0, 800);")
        time.sleep(scroll_pause_time)  # Ждем подгрузки новых элементов
        
        # Вычисляем новую высоту страницы
        new_height = driver.execute_script("return document.body.scrollHeight")
        
        # Если высота не изменилась, значит, достигнут конец страницы
        if new_height == last_height:
            # Дополнительная проверка: пробуем прокрутить еще раз с меньшей паузой
            driver.execute_script("window.scrollBy(0, 800);")
            time.sleep(scroll_pause_time / 2)  # Уменьшаем паузу для финальной проверки
            final_height = driver.execute_script("return document.body.scrollHeight")
            
            # Если высота снова не изменилась, завершаем прокрутку
            if final_height == new_height:
                logging.info("Достигнут конец страницы.")
                break  # Окончательно завершаем прокрутку
        
        # Обновляем высоту для следующей итерации
        last_height = new_height
def safe_find_element(driver, by, value, retries=3):
    for attempt in range(retries):
        try:
            return driver.find_element(by, value)
        except WebDriverException as e:
            logging.warning(f"Ошибка при поиске элемента {value}. Попытка {attempt + 1}/{retries}: {e}")
            time.sleep(2)  # Пауза перед повторной попыткой
    raise Exception(f"Не удалось найти элемент {value} после {retries} попыток.")

def save_to_excel(data, query):
    """
    Сохраняет данные в Excel-файл с форматированием для улучшения читаемости.
    :param data: Список словарей с данными о товарах.
    :param query: Поисковый запрос пользователя.
    """
    try:
        # Создаем "безопасное" имя файла, заменяя недопустимые символы
        safe_filename = re.sub(r'[\\/*?:"<>|]', "", query)  # Удаляем запрещенные символы
        safe_filename = safe_filename.replace(" ", "_")  # Заменяем пробелы на подчеркивания
        filename = f"{safe_filename}_products.xlsx"

        # Создаем DataFrame из списка словарей
        df = pd.DataFrame(data)
        # Переименовываем столбцы для удобства чтения
        df.rename(columns={
            "name": "Название товара",
            "price": "Цена",
            "rating_info": "Рейтинг",
            "link": "Ссылка на товар"
        }, inplace=True)

        # Создаем новую книгу Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Товары"

        # Добавляем заголовки
        headers = list(df.columns)
        for col_num, header in enumerate(headers, 1):
            col_letter = get_column_letter(col_num)
            cell = ws[f"{col_letter}1"]
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")  # Жирный белый текст
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")  # Синий фон
            cell.alignment = Alignment(horizontal="center", vertical="center")  # Выравнивание по центру
            cell.border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin")
            )

        # Добавляем данные
        for row_num, row_data in enumerate(data, 2):  # Начинаем со второй строки
            for col_num, value in enumerate(row_data.values(), 1):
                col_letter = get_column_letter(col_num)
                cell = ws[f"{col_letter}{row_num}"]
                cell.value = value
                cell.alignment = Alignment(horizontal="center", vertical="center")  # Выравнивание по центру
                cell.border = Border(
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                    top=Side(style="thin"),
                    bottom=Side(style="thin")
                )

        # Автоподбор ширины столбцов
        for col_num, column in enumerate(df.columns, 1):
            max_length = max(
                len(str(header)) for header in df[column]
            )  # Находим максимальную длину данных в столбце
            col_letter = get_column_letter(col_num)
            ws.column_dimensions[col_letter].width = max_length + 5  # Добавляем немного отступа

        # Сохраняем файл
        wb.save(filename)
        logging.info(f"Данные успешно сохранены в файл: {filename}")
        print(f"Данные успешно сохранены в файл: {filename}")
        return filename
    except Exception as e:
        logging.error(f"Ошибка при сохранении данных в Excel: {e}")
        print(f"Ошибка при сохранении данных в Excel: {e}")
        return None

def parse_wildberries(query: str):
    """
    Собирает данные из подклассов product-card__brand-wrap, product-card__price price, product-card__wrapper,
    product-card__rating-wrap и product-card__thermometer-wrap, используя прокрутку и пагинацию.
    :param query: Поисковый запрос пользователя.
    """
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # Запуск в полноразмерном окне
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    results = []  # Список для хранения результатов

    try:
        url = f"https://www.wildberries.ru/catalog/0/search.aspx?search={query}"
        logging.info(f"Перехожу на страницу: {url}")
        driver.get(url)

        wait = WebDriverWait(driver, 20)  # Увеличиваем время ожидания
        product_card_overflow = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "product-card-overflow")))

        page_number = 1
        while True:
            logging.info(f"Обработка страницы #{page_number}...")
            smooth_scroll_to_bottom(driver)

            middle_wraps = product_card_overflow.find_elements(By.CLASS_NAME, "product-card__middle-wrap")
            bottom_wraps = product_card_overflow.find_elements(By.CLASS_NAME, "product-card__bottom-wrap")

            total_cards = len(middle_wraps)
            logging.info(f"Найдено {total_cards} карточек товаров на странице #{page_number}.")

            for idx, wrap in enumerate(middle_wraps, start=1):
                try:
                    brand_wrap = wrap.find_element(By.CLASS_NAME, "product-card__brand-wrap")
                    name = brand_wrap.text.strip()

                    # Безопасный поиск элементов
                    price_text = safe_find_element(wrap, By.CLASS_NAME, "product-card__price.price").text.strip()
                    price_text = re.sub(r"с WB Кошельком", "", price_text).strip()
                    prices = re.split(r"₽", price_text)
                    formatted_prices = [f"{price.strip()}₽" for price in prices if price.strip()]
                    formatted_price = "\n".join(formatted_prices)

                    rating_text = safe_find_element(bottom_wraps[idx - 1], By.CLASS_NAME, "product-card__rating-wrap").text.strip()
                    link = safe_find_element(driver, By.CSS_SELECTOR, ".product-card__wrapper a.product-card__link.j-card-link.j-open-full-product-card").get_attribute("href")

                    results.append({
                        "name": name,
                        "price": formatted_price,
                        "rating_info": rating_text,
                        "link": link
                    })
                except Exception as e:
                    logging.warning(f"Ошибка при обработке карточки #{idx}: {e}")

            # Проверка кнопки "Следующая страница"
            try:
                next_page_button = driver.find_element(By.CLASS_NAME, "pagination-next.pagination__next.j-next-page")
                if "disabled" in next_page_button.get_attribute("class"):
                    break
                next_page_button.click()
                time.sleep(2)
                page_number += 1
            except Exception as e:
                logging.info("Кнопка 'Следующая страница' не найдена или произошла ошибка.")
                break

        print(f"Найдено {len(results)} карточек товаров.")
        return results

    except Exception as e:
        logging.error(f"Ошибка при парсинге: {e}")
        return results, None  # Возвращаем частично собранные данные
    finally:
        # Гарантированное сохранение данных в Excel
        if results:
            excel_filename = save_to_excel(results, query)
        driver.quit()  # Закрываем браузер
        return results, excel_filename
# Точка входа для запуска парсера напрямую
if __name__ == '__main__':
    # Запрос ввода от пользователя
    query = input("Введите название товара для поиска: ").strip()
    if not query:
        print("Пожалуйста, укажите название товара.")
    else:
        print("Выполняю поиск товаров...")
        parse_wildberries(query)