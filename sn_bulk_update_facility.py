
import os
import time
import csv
from dataclasses import dataclass
from typing import List, Tuple, Optional
from dotenv import load_dotenv

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Opcional: clique por coordenadas (PyAutoGUI para mouse/teclado)
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except Exception:
    PYAUTOGUI_AVAILABLE = False


@dataclass
class Config:
    instance_url: str
    username: str
    password: str
    excel_path: str
    excel_sheet: str
    excel_column: str
    facility_type_text: str = "Plant Location"
    use_coordinate_save: bool = False
    right_click_x: int = 1328
    right_click_y: int = 190
    implicit_wait_s: int = 2
    explicit_wait_s: int = 25
    max_rows: Optional[int] = None
    chrome_binary: Optional[str] = None
    use_isolated_profile: bool = True
    wait_before_search_s: int = 0
    # NOVOS CAMPOS para busca por coordenadas
    use_coordinate_search: bool = False
    search_click_x: int = 0
    search_click_y: int = 0


def project_path(*parts) -> str:
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)


def build_driver(cfg: Config) -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-infobars")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--window-size=1280,720")
    options.add_argument("--window-position=100,100")

    # Perfil isolado (mantém cookies/login sem afetar seu Chrome padrão)
    if cfg.use_isolated_profile:
        profile_dir = project_path("chrome-profile")
        os.makedirs(profile_dir, exist_ok=True)
        options.add_argument(f"--user-data-dir={profile_dir}")

    # Chrome portátil
    if cfg.chrome_binary:
        chrome_path = cfg.chrome_binary
        if not os.path.isabs(chrome_path):
            chrome_path = project_path(chrome_path)
        if not os.path.exists(chrome_path):
            raise FileNotFoundError(f"Chrome portátil não encontrado em: {chrome_path}")
        options.binary_location = chrome_path

    service = Service()
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def try_login(driver: webdriver.Chrome, cfg: Config):
    driver.get(cfg.instance_url)
    driver.implicitly_wait(cfg.implicit_wait_s)

    # Pausa manual para SSO/MFA (se habilitado no .env)
    if os.getenv("SSO_MODE", "false").lower() == "true":
        input("Conclua o login SSO/MFA no Chrome e pressione ENTER para continuar...")

    # Aguarda o carregamento completo após login/SSO
    try:
        WebDriverWait(driver, 60).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except Exception:
        pass

    # Se aparecer login clássico e você forneceu credenciais no .env, tenta login.
    try:
        user_field = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "user_name")))
        pass_field = driver.find_element(By.ID, "user_password")
        login_btn  = driver.find_element(By.ID, "sysverb_login")
        if cfg.username and cfg.password:
            user_field.clear(); user_field.send_keys(cfg.username)
            pass_field.clear(); pass_field.send_keys(cfg.password)
            login_btn.click()
            WebDriverWait(driver, 60).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
    except Exception:
        # Se não há login clássico, assumimos SSO concluído
        pass


def coordinate_type_and_search(cfg: Config, text: str):
    """
    Foca a caixa de busca global por coordenadas e digita o texto + ENTER usando PyAutoGUI.
    Requer USE_COORDINATE_SEARCH=true e SEARCH_CLICK_X/SEARCH_CLICK_Y configurados no .env.
    """
    if not PYAUTOGUI_AVAILABLE:
        raise RuntimeError("pyautogui não está disponível. Instale com: pip install pyautogui")

    if not cfg.use_coordinate_search:
        raise RuntimeError("USE_COORDINATE_SEARCH está falso. Habilite no .env para usar coordenadas na busca.")

    if cfg.search_click_x <= 0 or cfg.search_click_y <= 0:
        raise ValueError("Defina SEARCH_CLICK_X e SEARCH_CLICK_Y no .env para focar a caixa de busca.")

    # Garante que a janela ativa é o Chrome; recomenda-se mantê-lo maximizado
    time.sleep(0.5)
    pyautogui.moveTo(cfg.search_click_x, cfg.search_click_y, duration=0.25)
    pyautogui.click()
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(text, interval=0.02)
    pyautogui.press('enter')


def search_value(driver: webdriver.Chrome, cfg: Config, value: str):
    """
    Em vez de usar seletores do HTML para a busca global, foca a caixa por coordenadas
    e digita com PyAutoGUI. Depois, espera a página de resultado carregar.
    """
    # Espera opcional (SSO/JS da home)
    if cfg.wait_before_search_s > 0:
        time.sleep(cfg.wait_before_search_s)

    # Garante carregamento geral
    try:
        WebDriverWait(driver, 60).until(lambda d: d.execute_script("return document.readyState") == "complete")
    except Exception:
        pass

    # Foca e digita usando coordenadas
    coordinate_type_and_search(cfg, value)

    # Aguarda algum resultado (lista/form/iframe) — via Selenium
    wait = WebDriverWait(driver, max(cfg.explicit_wait_s, 45))
    time.sleep(1)  # transição breve
    wait.until(EC.any_of(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table.list_table, .list2_body")),
        EC.presence_of_element_located((By.CSS_SELECTOR, "form")),
        EC.presence_of_element_located((By.CSS_SELECTOR, "#gsft_main"))
    ))


def open_record_for_value(driver: webdriver.Chrome, cfg: Config, value: str):
    """Abre o registro correspondente ao texto buscado (hostname)."""
    wait = WebDriverWait(driver, max(cfg.explicit_wait_s, 45))

    # 1) Link com texto exato
    try:
        link = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[normalize-space(text())='{value}']")))
        link.click(); return
    except TimeoutException:
        pass

    # 2) Célula com texto exato -> clicar no primeiro link da linha
    try:
        cell = wait.until(EC.element_to_be_clickable((By.XPATH, f"//td[normalize-space()='{value}']")))
        row = cell.find_element(By.XPATH, "./ancestor::tr")
        row_link_candidates = row.find_elements(By.XPATH, ".//a[not(@aria-hidden='true')]")
        if row_link_candidates:
            row_link_candidates[0].click(); return
    except TimeoutException:
        pass

    # 3) Se já estamos no formulário, segue
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "form")))
    except TimeoutException:
        raise TimeoutException(f"Não foi possível abrir o registro para '{value}'.")


def set_facility_type(driver: webdriver.Chrome, cfg: Config):
    wait = WebDriverWait(driver, max(cfg.explicit_wait_s, 45))
    try:
        elem = wait.until(EC.presence_of_element_located((By.ID, "cmdb_ci_computer.u_facility_type")))
        tag = elem.tag_name.lower()
        if tag == "select":
            Select(elem).select_by_visible_text(cfg.facility_type_text)
        else:
            elem.clear(); elem.send_keys(cfg.facility_type_text); elem.send_keys(Keys.TAB)
        return
    except TimeoutException:
        pass
    try:
        label = wait.until(EC.presence_of_element_located((By.XPATH, "//label[normalize-space()='Facility type' or contains(., 'Facility type')]")))
        container = label.find_element(By.XPATH, "./ancestor::*[self::div or self::td]")
        field = container.find_element(By.XPATH, ".//select|.//input|.//textarea")
        if field.tag_name.lower() == "select":
            Select(field).select_by_visible_text(cfg.facility_type_text)
        else:
            field.clear(); field.send_keys(cfg.facility_type_text); field.send_keys(Keys.TAB)
    except Exception as e:
        raise Exception(f"Não foi possível localizar/definir Facility type: {e}")


def save_record_via_dom(driver: webdriver.Chrome, cfg: Config):
    wait = WebDriverWait(driver, max(cfg.explicit_wait_s, 45))
    for candidate in ["sysverb_save", "sysverb_update", "save_button", "update_button"]:
        try:
            btn = wait.until(EC.element_to_be_clickable((By.ID, candidate)))
            btn.click(); return
        except TimeoutException:
            continue
    try:
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Save' or normalize-space()='Update']")))
        btn.click(); return
    except TimeoutException:
        pass
    try:
        actions_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@aria-label,'More') or contains(@class,'btn') and (contains(.,'More') or contains(.,'Actions'))]")))
        actions_btn.click()
        menu_item = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@class,'menu') or contains(@role,'menu')]//*[normalize-space()='Save' or normalize-space()='Update']")))
        menu_item.click(); return
    except TimeoutException:
        raise TimeoutException("Não foi possível encontrar Save/Update via DOM.")


def save_record_via_coordinates(driver: webdriver.Chrome, cfg: Config):
    if not PYAUTOGUI_AVAILABLE:
        raise RuntimeError("pyautogui não está disponível. Instale com: pip install pyautogui")
    driver.maximize_window(); time.sleep(0.7)
    pyautogui.moveTo(cfg.right_click_x, cfg.right_click_y, duration=0.25)
    pyautogui.click(button="right"); time.sleep(0.4)
    wait = WebDriverWait(driver, max(cfg.explicit_wait_s, 45))
    try:
        menu_save = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@class,'menu') or contains(@role,'menu')]//*[normalize-space()='Save']")))
        menu_save.click()
    except TimeoutException:
        pyautogui.press('down'); pyautogui.press('enter')


def read_excel(cfg: Config) -> List[str]:
    df = pd.read_excel(project_path(cfg.excel_path), sheet_name=cfg.excel_sheet, engine="openpyxl")
    if cfg.excel_column not in df.columns:
        raise ValueError(f"Coluna '{cfg.excel_column}' não encontrada na aba '{cfg.excel_sheet}'.")
    series = df[cfg.excel_column].astype(str).str.strip()
    values = [v for v in series.tolist() if v and v.lower() not in ("nan", "none")]
    seen = set(); dedup = []
    for v in values:
        if v not in seen:
            dedup.append(v); seen.add(v)
    if cfg.max_rows:
        dedup = dedup[:cfg.max_rows]
    return dedup


def process_item(driver: webdriver.Chrome, cfg: Config, value: str) -> Tuple[str, str]:
    try:
        search_value(driver, cfg, value)
        open_record_for_value(driver, cfg, value)
        set_facility_type(driver, cfg)
        if cfg.use_coordinate_save:
            save_record_via_coordinates(driver, cfg)
        else:
            save_record_via_dom(driver, cfg)
        time.sleep(1.0)
        return ("OK", "Atualizado e salvo.")
    except TimeoutException as te:
        try:
            driver.save_screenshot(project_path(f"error_{value}.png"))
        except Exception:
            pass
        return ("ERROR", f"Timeout: {te}")
    except Exception as e:
        try:
            driver.save_screenshot(project_path(f"error_{value}.png"))
        except Exception:
            pass
        return ("ERROR", f"{type(e).__name__}: {e}")


def write_report(rows: List[Tuple[str, str, str]], path: str = "resultado_facility.csv"):
    out_path = project_path(path)
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["HOSTNAME", "STATUS", "DETALHE"])
        writer.writerows(rows)
    print(f"Relatorio salvo em: {out_path}")


def main():
    load_dotenv(project_path(".env"))
    cfg = Config(
        instance_url=os.getenv("INSTANCE_URL", "").strip(),
        username=os.getenv("SN_USER", "").strip(),
        password=os.getenv("SN_PASS", "").strip(),
        excel_path=os.getenv("EXCEL_PATH", "Inventario_RAD.xlsx").strip(),
        excel_sheet=os.getenv("EXCEL_SHEET", "INVENTARIO RAD").strip(),
        excel_column=os.getenv("EXCEL_COLUMN", "HOSTNAME").strip(),
        facility_type_text=os.getenv("FACILITY_TYPE", "Plant Location").strip(),
        use_coordinate_save=os.getenv("USE_COORDINATE_SAVE", "false").lower() == "true",
        right_click_x=int(os.getenv("RIGHT_CLICK_X", "1328")),
        right_click_y=int(os.getenv("RIGHT_CLICK_Y", "190")),
        max_rows=int(os.getenv("MAX_ROWS", "0")) or None,
        chrome_binary=os.getenv("CHROME_BINARY", "").strip() or None,
        wait_before_search_s=int(os.getenv("WAIT_BEFORE_SEARCH", "0")),
        use_coordinate_search=os.getenv("USE_COORDINATE_SEARCH", "false").lower() == "true",
        search_click_x=int(os.getenv("SEARCH_CLICK_X", "0")),
        search_click_y=int(os.getenv("SEARCH_CLICK_Y", "0")),
    )

    if not cfg.instance_url or not cfg.excel_path:
        raise ValueError("INSTANCE_URL e EXCEL_PATH precisam estar definidos no .env.")

    values = read_excel(cfg)
    print(f"Itens a processar (HOSTNAME): {len(values)}" + (f" (limitado a {cfg.max_rows})" if cfg.max_rows else ""))

    driver = build_driver(cfg)
    rows = []
    try:
        try_login(driver, cfg)
        for i, value in enumerate(values, start=1):
            print(f"[{i}/{len(values)}] {value}")
            status, detail = process_item(driver, cfg, value)
            print(f"   -> {status}: {detail}")
            rows.append((value, status, detail))
            time.sleep(0.4)
        write_report(rows)
        print("Processo concluido.")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
