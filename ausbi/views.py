import os
import time
from selenium import webdriver
from django.shortcuts import render
from django.http import HttpResponse
from .forms import MesAnoForm
from .models import CentroCusto
from selenium.webdriver.common.by import By
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def configure_download_folder(download_folder):
    firefox_options = webdriver.FirefoxOptions()
    firefox_options.set_preference("browser.download.folderList", 2)
    firefox_options.set_preference("browser.download.manager.showWhenStarting", False)
    firefox_options.set_preference("browser.download.dir", download_folder)
    firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    driver = webdriver.Firefox(options=firefox_options)
    return driver






def login(driver, usuario, senha):
    driver.get("https://app.folhacerta.com/Admin/Login")
    time.sleep(5)
    
    # Preenchendo o campo de usuário usando JavaScript
    driver.execute_script(f"document.getElementsByName('login')[0].value = '{usuario}'")
    
    # Preenchendo o campo de senha
    campo_senha = driver.find_element(By.NAME, "senha")
    campo_senha.send_keys(senha)
    
    # Clicando no botão de login
    botao_login = driver.find_element(By.XPATH, "//button[text()='Login']")
    botao_login.click()
    
    time.sleep(5)
    input("Pressione Enter após resolver o captcha manualmente...")




    

def baixar_planilhas(request):
    if request.method == 'POST':
        form = MesAnoForm(request.POST)
        if form.is_valid():
            ano = form.cleaned_data['ano']
            mes = form.cleaned_data['mes']
            centros_custo = CentroCusto.objects.all()
            os.makedirs('planilhas', exist_ok=True)
            download_folder = r"C:\Users\alexs\Downloads"
            driver = configure_download_folder(download_folder)

            login(driver, "40997778806", "2241")

            for centro in centros_custo:
                centro_custo_id = centro.centro_custo_id
                baixou = False
                tentativas = 0

                while not baixou and tentativas < 3:
                    url = f"https://app.folhacerta.com/Admin/HorasTrabalhadasDiaExcel?ano={ano}&mes={mes}&centrocusto_id={centro_custo_id}&filtro=&usuario=99505"
                    driver.get(url)
                    filename = os.path.join(download_folder, f"planilha_{ano}_{mes}_{centro_custo_id}.xlsx")
                    baixou = aguardar_download(filename, 300)
                    if not baixou:
                        time.sleep(5)
                    tentativas += 1

            driver.quit()
            return HttpResponse("Planilhas baixadas com sucesso.")

    else:
        form = MesAnoForm()

    return render(request, 'robotic/relatorio_dia_fc.html', {'form': form})


def aguardar_download(filename, timeout):
    for _ in range(timeout):
        if os.path.exists(filename) and os.path.getsize(filename) > 0:
            return True
        time.sleep(1)
    return False
