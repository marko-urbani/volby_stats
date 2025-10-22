import os
import time
import pandas as pd
import numpy as np
from io import StringIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Cílová složka pro uložení finálních CSV souborů
TARGET_DIR = '.' 

# Základní URL pro výsledky voleb 2025
BASE_URL = 'https://www.volby.cz/app/ps2025/cs/results'

# Slovník krajů s jejich kódy a názvy pro soubory
KRAJE = {
    'Celkem': '',
    'Hlavni_mesto_Praha': '1100',
    'Stredocesky_kraj': '2100',
    'Jihocesky_kraj': '3100',
    'Plzensky_kraj': '3200',
    'Karlovarsky_kraj': '4100',
    'Ustecky_kraj': '4200',
    'Liberecky_kraj': '5100',
    'Kralovehradecky_kraj': '5200',
    'Pardubicky_kraj': '5300',
    'Kraj_Vysocina': '6100',
    'Jihomoravsky_kraj': '6200',
    'Olomoucky_kraj': '7100',
    'Zlinsky_kraj': '7200',
    'Moravskoslezsky_kraj': '8100',
    'Zahranici': '9900'
}


def scrape_clean_and_save(nazev, kod):
    """
    Načte stránku, inteligentně zpracuje data podle struktury (kraj vs. zahraničí)
    a uloží finální čisté CSV.
    """
    if kod:
        url = f'{BASE_URL}/!___{kod}'
    else:
        url = BASE_URL
    
    nazev_souboru = f'vysledky_ps2025_{nazev}.csv'
    cesta_k_souboru = os.path.join(TARGET_DIR, nazev_souboru)

    print(f"Zpracovávám: {nazev}...")

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    try:
        driver.get(url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr")))
        time.sleep(1) 

        html_source = driver.page_source
        
        df_raw = pd.read_html(
            StringIO(html_source), 
            decimal=',',
            thousands='\xa0'
        )[1]

        # --- KLÍČOVÁ ZMĚNA: Ošetření pro zahraničí ---
        is_zahranici = (nazev == 'Zahranici')

        if is_zahranici:
            # Zahraničí má méně sloupců (chybí mandáty)
            df = pd.DataFrame({
                'nazev_strany': df_raw.iloc[:, 1],
                'hlasy_celkem': df_raw.iloc[:, 2],
                'hlasy_procenta': df_raw.iloc[:, 3],
            })
            # Doplníme chybějící sloupce nulami pro konzistentní strukturu
            df['mandaty_pocet'] = 0
            df['mandaty_procenta'] = 0.0
        else:
            # Standardní zpracování pro kraje
            if df_raw.shape[1] < 6:
                raise ValueError("Tabulka pro kraj nemá očekávaný počet sloupců.")
            df = pd.DataFrame({
                'nazev_strany': df_raw.iloc[:, 1],
                'hlasy_celkem': df_raw.iloc[:, 2],
                'hlasy_procenta': df_raw.iloc[:, 3],
                'mandaty_pocet': df_raw.iloc[:, 4],
                'mandaty_procenta': df_raw.iloc[:, 5]
            })
        
        # Finální úpravy datových typů
        df = df.dropna(subset=['hlasy_celkem'])
        df['hlasy_celkem'] = df['hlasy_celkem'].astype(np.int64)
        df['mandaty_pocet'] = df['mandaty_pocet'].fillna(0).astype(int)

        df.to_csv(cesta_k_souboru, index=False, encoding='utf-8-sig', sep=';', decimal='.')
        print(f'✔ Hotovo. Data pro {nazev} uložena do: {cesta_k_souboru}')

    except IndexError:
        print(f"❌ Chyba: Pro {nazev} se nepodařilo najít očekávanou tabulku výsledků na stránce.")
    except Exception as e:
        print(f'❌ Nastala neočekávaná chyba pro {nazev}: {e}')
    finally:
        driver.quit()


if __name__ == '__main__':
    print("--- Zahajuji stahování a čištění kompletních výsledků voleb 2025 ---")
    os.makedirs(TARGET_DIR, exist_ok=True)
    
    for nazev_kraje, kod_kraje in KRAJE.items():
        scrape_clean_and_save(nazev_kraje, kod_kraje)
        
    print("\n--- Všechny operace dokončeny. ---")