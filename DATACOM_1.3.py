from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from datetime import datetime, timedelta
import os


def initialize_driver(driver_path):
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-webgl")

    service = Service(driver_path)
    return webdriver.Chrome(service=service, options=chrome_options)


def close_overlays(driver):
    try:
        popup_close_button = driver.find_element(By.CLASS_NAME, "popupCloseIcon")
        if popup_close_button.is_displayed():
            popup_close_button.click()
            print("Overlay fechado com sucesso!")
    except Exception:
        print("Nenhum overlay encontrado.")


def click_with_retry(driver, element):
    for _ in range(3):  # Tentar clicar até 3 vezes
        try:
            element.click()
            return
        except Exception:
            print("Clique direto falhou, tentando via JavaScript.")
            try:
                driver.execute_script("arguments[0].click();", element)
                return
            except Exception as e:
                print(f"Erro ao clicar no elemento: {e}")
                time.sleep(1)  # Esperar antes de tentar novamente


def select_countries(driver, wait):
    try:
        filter_button = wait.until(EC.element_to_be_clickable((By.ID, "filterStateAnchor")))
        filter_button.click()
        print("Botão de filtros clicado para abrir os filtros!")

        wait.until(EC.presence_of_element_located((By.ID, "calendarFilterBox_country")))
        print("Lista de países carregada com sucesso!")

        country_checkboxes = driver.find_elements(By.XPATH, '//input[@name="country[]"]')
        for checkbox in country_checkboxes:
            country_id = checkbox.get_attribute("id")
            if checkbox.is_selected() and country_id != "country32":
                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                    time.sleep(0.5)
                    checkbox.click()
                    print(f"Desmarcado: {country_id}")
                except Exception:
                    print(f"Falha ao clicar no checkbox {country_id}, tentando com JavaScript.")
                    driver.execute_script("arguments[0].click();", checkbox)

        brazil_checkbox = driver.find_element(By.ID, "country32")
        if not brazil_checkbox.is_selected():
            driver.execute_script("arguments[0].scrollIntoView(true);", brazil_checkbox)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", brazil_checkbox)
        print("Brasil selecionado com sucesso!")

        apply_button = driver.find_element(By.ID, "ecSubmitButton")
        driver.execute_script("arguments[0].scrollIntoView(true);", apply_button)
        time.sleep(0.5)
        click_with_retry(driver, apply_button)
        print("Filtros aplicados com sucesso!")
    except Exception as e:
        print(f"Erro ao selecionar apenas o Brasil: {e}")


def select_dates(driver, wait, start_date, end_date):
    try:
        date_picker_button = wait.until(EC.element_to_be_clickable((By.ID, "datePickerToggleBtn")))
        driver.execute_script("arguments[0].scrollIntoView(true);", date_picker_button)
        time.sleep(0.5)
        click_with_retry(driver, date_picker_button)
        print("Botão do seletor de datas clicado!")

        wait.until(EC.presence_of_element_located((By.ID, "ui-datepicker-div")))
        print("Seletor de datas carregado com sucesso!")

        start_date_input = driver.find_element(By.ID, "startDate")
        start_date_input.clear()
        start_date_input.send_keys(start_date)
        print(f"Data de início {start_date} selecionada com sucesso!")

        end_date_input = driver.find_element(By.ID, "endDate")
        end_date_input.clear()
        end_date_input.send_keys(end_date)
        print(f"Data de término {end_date} selecionada com sucesso!")

        apply_button = driver.find_element(By.ID, "applyBtn")
        driver.execute_script("arguments[0].scrollIntoView(true);", apply_button)
        time.sleep(0.5)
        close_overlays(driver)
        click_with_retry(driver, apply_button)
        print("Botão 'Aplicar' clicado com sucesso!")
        time.sleep(20)
    except Exception as e:
        print(f"Erro ao selecionar as datas: {e}")


def extract_table_data(driver, wait):
    try:
        wait.until(EC.presence_of_element_located((By.ID, "dividendsCalendarData")))
        print("Tabela carregada com sucesso!")

        table = driver.find_element(By.ID, "dividendsCalendarData")
        rows = table.find_elements(By.TAG_NAME, "tr")

        data = []
        for row in rows:
            if "theDay" in row.get_attribute("class"):
                continue

            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) == 7:
                data.append({
                    "Empresa": cells[1].text.strip(),
                    "Data ex-dividendos": cells[2].text.strip(),
                    "Dividendo": cells[3].text.strip(),
                    "Tipo": cells[4].text.strip(),
                    "Pagamento": cells[5].text.strip(),
                    "Rendimento": cells[6].text.strip(),
                })

        df = pd.DataFrame(data)
        print("Dados extraídos com sucesso!")
        return df

    except Exception as e:
        print(f"Erro ao extrair os dados da tabela: {e}")
        return pd.DataFrame()


# Configurações Gerais
driver_path = r'C:\Users\João Vitor\OneDrive\Worksplace Python\WebDriver\chromedriver.exe'
output_path = r"C:\Users\João Vitor\OneDrive\Projeto de investimento\Datacom Proventos"
os.makedirs(output_path, exist_ok=True)

current_date = datetime(2024, 1, 1)
end_date = datetime.now()

while current_date <= end_date:
    try:
        start_date = current_date.strftime("%d/%m/%Y")
        last_day_of_month = (current_date.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        end_date_str = last_day_of_month.strftime("%d/%m/%Y")

        print(f"Extraindo dados para o período: {start_date} a {end_date_str}")

        # Inicializar o driver
        driver = initialize_driver(driver_path)
        wait = WebDriverWait(driver, 15)

        # Navegar até a página e configurar os filtros
        url = "https://br.investing.com/dividends-calendar/"
        driver.get(url)
        close_overlays(driver)
        select_countries(driver, wait)
        select_dates(driver, wait, start_date=start_date, end_date=end_date_str)

        # Extrair os dados
        df = extract_table_data(driver, wait)
        if not df.empty:
            filename = f"DATACOM_{end_date_str.replace('/', '-')}.csv"
            filepath = os.path.join(output_path, filename)
            df.to_csv(filepath, index=False, encoding="utf-8")
            print(f"Dados salvos em {filepath}.")

    except Exception as e:
        print(f"Erro durante a extração para o período {start_date} a {end_date_str}: {e}")

    finally:
        # Fechar o driver após cada mês
        driver.quit()
        current_date = last_day_of_month + timedelta(days=1)
