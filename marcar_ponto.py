from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import TimeoutException
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
import os
import time

# Horários válidos para marcação
horarios_validos = ["11:00", "16:00", "17:00", "20:00"]

# Carrega variáveis do arquivo .env
load_dotenv()
usuario = os.getenv("USUARIO")
senha = os.getenv("SENHA")

if not usuario or not senha:
    print("❌ Erro: Credenciais não encontradas no arquivo .env")
    exit(1)

# Configuração do navegador
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)

# Caminho do ChromeDriver local
chromedriver_path = Path(__file__).parent / "chromedriver-win64" / "chromedriver-win64" / "chromedriver.exe"
service = Service(str(chromedriver_path))

try:
    driver = webdriver.Chrome(service=service, options=chrome_options)
except Exception as e:
    print(f"⚠️ Aviso: {e}")
    print("Tentando continuar mesmo com incompatibilidade de versão...")
    driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    print("🔄 Acessando site de ponto...")
    driver.get("https://newponto.positivosmais.com/webponto/default.asp")

    # Aguarda os campos de login
    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "requiredusuario")))
    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "requiredsenha")))

    print("✍️ Preenchendo usuário e senha...")
    campo_usuario = driver.find_element(By.ID, "requiredusuario")
    campo_senha = driver.find_element(By.ID, "requiredsenha")

    campo_usuario.clear()
    campo_usuario.send_keys(usuario)
    campo_senha.clear()
    campo_senha.send_keys(senha)

    # Clica no botão OK
    botao_ok = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and contains(@value, 'OK')]"))
    )
    botao_ok.click()
    print("✅ Login realizado com sucesso.")

    # Trata alerta inesperado
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = Alert(driver)
        print("⚠️ Alerta detectado:", alert.text)
        alert.accept()
    except:
        print("✅ Nenhum alerta detectado.")

    # Aguarda redirecionamento
    time.sleep(5)
    print("🔗 URL após login:", driver.current_url)

    # Clica em "Lançamentos"
    menu_lancamentos = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "menu2"))
    )
    menu_lancamentos.click()
    print("🟢 Clicou em 'Lançamentos'")

    # Clica em "Marcação do Ponto"
    marcacao_ponto = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='menu2_Item3']//a[contains(text(), 'Marcação do Ponto')]"))
    )
    marcacao_ponto.click()
    print("🟢 Clicou em 'Marcação do Ponto'")

    # Aguarda nova janela abrir e troca para ela
    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])
    print("🔄 Trocou para a nova janela de marcação")

    # Aguarda o campo de hora aparecer
    campo_hora = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@class='CampoCentro' and @type='text' and contains(@value, ':')]"))
    )
    hora_valor_raw = campo_hora.get_attribute("value")
    hora_valor = hora_valor_raw.strip() if hora_valor_raw else ""
    print(f"🕒 Horário exibido na página: {hora_valor}")

    # Verifica se é dia útil
    agora = datetime.now()
    dia_semana = agora.weekday()  # 0 = segunda, 6 = domingo

    if dia_semana < 5:
        if hora_valor in horarios_validos:
            print("✅ Horário válido! Marcando o ponto...")
            try:
                botao_ok = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, "Button1"))
                )
                botao_ok.click()
                print("🟢 Ponto marcado com sucesso.")
            except TimeoutException:
                print("⛔ Botão OK não encontrado ou não clicável.")
        else:
            print(f"⛔ Horário da página ({hora_valor}) não está na lista de marcação. Nenhuma ação será tomada.")
    else:
        print("⛔ Hoje não é dia útil. Nenhuma marcação será feita.")

    # Tempo para visualizar
    time.sleep(30)

except Exception as e:
    print("❌ Ocorreu um erro durante a execução:")
    print(str(e))
    time.sleep(30)

finally:
    print("^ Encerrando navegador.")
    driver.quit()
