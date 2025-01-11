from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import time
import random
import pyperclip
import subprocess
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# ------------------------------------------------------------------------------------

# Lista de CNPJs para processamento
cnpj_list = ["00000000000000"] #Aqui voce coloca os CNPJs que quer consultar

# ------------------------------------------------------------------------------------

def espera_aleatoria():
    time.sleep(random.uniform(1.5, 5))
    
def abrir_chrome_modo_depuracao_com_site(porta=9222, url="https://cav.receita.fazenda.gov.br/autenticacao/login"):
    # Caminho padrão para o Google Chrome, ajuste se necessário
    caminho_chrome = "Caminho/do/chrome.exe"
    
    # Parâmetros para abrir o Chrome em modo de depuração
    comando = [
        caminho_chrome,
        f"--remote-debugging-port={porta}",
        "--user-data-dir=C:\\chrome_debug_profile",
        url  # Adiciona a URL ao comando
    ]
    
    try:
        subprocess.Popen(comando)
        print(f"Chrome aberto em modo de depuração na porta {porta}. Site: {url}")
    except FileNotFoundError:
        print("Erro: Caminho do Google Chrome não encontrado. Verifique o caminho no código.")
    except Exception as e:
        print(f"Ocorreu um erro ao abrir o Chrome: {e}")

# ------------------------------------------------------------------------------------

def clicar_botao(x, y, descricao):
    espera_aleatoria()
    # Move o mouse para o botão e clica
    pyautogui.moveTo(x, y, duration=1)
    pyautogui.click()
    print(f"Botão '{descricao}' clicado com sucesso!")

def EntrarcomGov():
    espera_aleatoria()
    # Coordenadas do botão "Entrar com Gov"
    x = 885  # Substitua pelo valor de X encontrado
    y = 479  # Substitua pelo valor de Y encontrado
    espera_aleatoria()
    clicar_botao(x, y, "Entrar com Gov")

def SeuCertificadoDigital():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 945  # Substitua pelo valor de X encontrado
    y = 647  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "Seu Certificado Digital")
    
def OkCertificado():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 799  # Substitua pelo valor de X encontrado
    y = 357  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "OkCertificado")
    
def AlterarPerfil():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 1080  # Substitua pelo valor de X encontrado
    y = 219  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "AlterarPerfil")
    
def ClicarCNPJ():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 542  # Substitua pelo valor de X encontrado
    y = 381  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "ClicarCNPJ")
    
def SituacaoFiscal():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 480  # Substitua pelo valor de X encontrado
    y = 270  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "SituacaoFiscal")
    
def ConsultaPendencias():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 305  # Substitua pelo valor de X encontrado
    y = 489  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "ConsultaPendencias")
    
def GerarRelatorio():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 86  # Substitua pelo valor de X encontrado
    y = 332  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "GerarRelatorio")
    
def BaixarRelatorio():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 874  # Substitua pelo valor de X encontrado
    y = 384  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "BaixarRelatorio")
    
def PDFBaixado():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 896  # Substitua pelo valor de X encontrado
    y = 533  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "PDFBaixado")
    
def DiagnosticoFiscal():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 101  # Substitua pelo valor de X encontrado
    y = 292  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "DiagnosticoFiscal")
    
def EmitirCertidao():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 485  # Substitua pelo valor de X encontrado
    y = 445  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "EmitirCertidao")
    
def PDFBaixado2():
    espera_aleatoria()
    # Coordenadas do botão "Seu Certificado Digital"
    x = 894  # Substitua pelo valor de X encontrado
    y = 578  # Substitua pelo valor de Y encontrado
    espera_aleatoria()  # Aguarda a interface carregar antes de clicar
    clicar_botao(x, y, "PDFBaixadoCertidao")

def ColarCNPJ():
    espera_aleatoria()
    try:
        # Verifica se há CNPJs na lista
        if not cnpj_list:
            print("Nenhum CNPJ disponível na lista para colar.")
            return

        # Copiar o primeiro CNPJ da lista para a área de transferência
        cnpj = cnpj_list.pop(0)  # Remove e retorna o primeiro CNPJ
        pyperclip.copy(cnpj)

        # Simula colagem do CNPJ no campo (Ctrl + V)
        espera_aleatoria()
        pyautogui.hotkey("ctrl", "v")
        espera_aleatoria()

        print(f"CNPJ {cnpj} colado com sucesso!")
    except Exception as e:
        print(f"Erro em ColarCNPJ: {e}")
        
def ClicarAlterar():
    espera_aleatoria()  # Aguarda um tempo aleatório para parecer mais humano
    try:
        # Simula o pressionamento da tecla Tab para navegar
        pyautogui.press("tab")
        espera_aleatoria()

        # Simula o pressionamento da tecla Enter para ativar o botão
        pyautogui.press("enter")
        espera_aleatoria()

        print("Ação 'ClicarAlterar' realizada com sucesso usando Tab e Enter!")
    
    except Exception as e:
        print(f"Erro ao executar 'ClicarAlterar': {e}")
        
def FecharJanela():
    try:
        # Limpar o cache do navegador via teclas de atalho (Ctrl + Shift + Del)
        espera_aleatoria()
        pyautogui.hotkey('ctrl', 'shift', 'del')
        espera_aleatoria()
        
        # Aguarda a interface abrir e pressiona Enter para confirmar
        pyautogui.press('enter')
        espera_aleatoria()

        # Fecha o navegador
        pyautogui.hotkey('alt', 'f4')
        print("Cache limpo e navegador fechado com sucesso.")
    except Exception as e:
        print(f"Erro ao tentar limpar cache e fechar navegador: {e}")
    
# ------------------------------------------------------------------------------------
        
def TelaInicial():
    try:
        # Localizar o iframe, se necessário
        #iframe = driver.find_element(By.TAG_NAME, "iframe")  # Substitua por um localizador específico do iframe, se necessário
        #driver.switch_to.frame(iframe)

        # Aguarde até que o elemento com o texto "RECEITA FEDERAL DO BRASIL" esteja visível
        wait = WebDriverWait(driver, 10)  # Aguarde até 10 segundos
        elemento = wait.until(EC.presence_of_element_located((By.ID, "logoEcac")))

        # Clicar no elemento
        elemento.click()
        print("TelaInicial clicado com sucesso!")

        # Voltar para o contexto principal
        #driver.switch_to.default_content()

    except Exception as e:
        print("Erro encontrado:", e)

    finally:
        # Não fechar o navegador automaticamente para continuar na aba
        pass
        
# ------------------------------------------------------------------------------------

if __name__ == "__main__":
    start_time = time.time()  # Marca o tempo de início
    # Loop para processar todos os CNPJs
    while cnpj_list:  # Continua enquanto houver CNPJs na lista
        try:
            # Executa as ações necessárias para um CNPJ
            abrir_chrome_modo_depuracao_com_site()
            EntrarcomGov()
            SeuCertificadoDigital()
            OkCertificado()
            AlterarPerfil()
            ClicarCNPJ()
            ColarCNPJ()
            ClicarAlterar()
            SituacaoFiscal()
            ConsultaPendencias()
            time.sleep(10)
            GerarRelatorio()
            espera_aleatoria()
            BaixarRelatorio()
            time.sleep(10)
            PDFBaixado()
            espera_aleatoria()
            DiagnosticoFiscal()
            espera_aleatoria()
            EmitirCertidao()
            time.sleep(10)
            PDFBaixado2()
            
            # Configuração do WebDriver em modo de depuração
            chrome_options = Options()
            chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")  # Porta para depuração

            driver_path = 'Caminho/do/chrome/driver'  # Substitua pelo caminho do seu WebDriver
            service = Service(driver_path)

            # Conectando ao navegador já aberto
            driver = webdriver.Chrome(service=service, options=chrome_options)

            TelaInicial()
            FecharJanela()
            
            print("CNPJ processado com sucesso. Continuando para o próximo...")
        except Exception as e:
            print(f"Erro no processamento do CNPJ: {e}")
            break  # Encerra em caso de erro grave
        
    print("Todas os relatórios foram baixados!")  
    end_time = time.time()  # Marca o tempo de término
    duration = end_time - start_time  # Calcula o tempo de execução em segundos
    print(f"O código foi executado por {duration / 60:.0f} minutos.")   
