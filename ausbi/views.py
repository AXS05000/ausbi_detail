import os
import time
from .models import CentroCusto
from .forms import MesAnoForm
from django.shortcuts import render
from django.http import HttpResponse
from selenium import webdriver


def configure_download_folder(download_folder):
    # Opções para configurar o navegador Firefox
    firefox_options = webdriver.FirefoxOptions()
    
    # Configurar o diretório de download
    firefox_options.set_preference("browser.download.folderList", 2)
    firefox_options.set_preference("browser.download.manager.showWhenStarting", False)
    firefox_options.set_preference("browser.download.dir", download_folder)
    firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # Iniciar o navegador com as opções configuradas
    driver = webdriver.Firefox(options=firefox_options)
    
    return driver


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

            # Armazene o identificador da janela original
            original_window = driver.current_window_handle

            for centro in centros_custo:
                centro_custo_id = centro.centro_custo_id
                baixou = False
                tentativas = 0

                while not baixou and tentativas < 3:  # tente baixar até 3 vezes
                    # Use JavaScript para abrir uma nova aba
                    driver.execute_script("window.open('about:blank', '_blank');")
                    # Mude para a nova aba
                    driver.switch_to.window(driver.window_handles[1])
                    
                    url = f"https://app.folhacerta.com/Admin/HorasTrabalhadasDiaExcel?ano={ano}&mes={mes}&centrocusto_id={centro_custo_id}&filtro=&usuario=99505"
                    driver.get(url)  # Navegue até a URL usando Selenium.

                    filename = os.path.join(download_folder, f"planilha_{ano}_{mes}_{centro_custo_id}.xlsx")

                    # Aguarde até que o arquivo seja baixado.
                    baixou = aguardar_download(filename, 300)  # Espera até 5 minutos. Ajuste conforme necessário.

                    if not baixou:
                        time.sleep(5)  # Aguarde 5 segundos antes de tentar novamente.
                    tentativas += 1

                    # Feche a aba atual
                    driver.close()

                    # Volte para a aba original
                    driver.switch_to.window(original_window)

            # Finalmente, feche o navegador após o download de todos os arquivos.
            driver.quit()
            return HttpResponse("Planilhas baixadas com sucesso.")

    else:
        form = MesAnoForm()

    return render(request, 'robotic/relatorio_dia_fc.html', {'form': form})


def aguardar_download(filename, timeout):
    """
    Aguarde até que o arquivo seja baixado ou o tempo limite seja atingido.
    """
    for _ in range(timeout):
        if os.path.exists(filename) and os.path.getsize(filename) > 0:
            return True
        time.sleep(1)
    return False
