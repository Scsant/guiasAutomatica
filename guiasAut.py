import re
import os
import time
import tempfile
from io import BytesIO
import shutil
import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tempfile
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import chromedriver_autoinstaller

import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def get_webdriver():
    options = Options()

    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920x1080")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--single-process")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-default-apps")
    options.add_argument("--disable-sync")
    options.add_argument("--metrics-recording-only")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")

    # Caminho para o chromedriver incluído no repositório
    chromedriver_path = os.path.join(os.getcwd(), "chromedriver")
    service = Service(executable_path=chromedriver_path)

    return webdriver.Chrome(service=service, options=options)






# Função para executar o Selenium
def run_selenium(logpath: str, url: str, numero_doc_input: str, valor_input: str, chave_nf_input: str) -> None:

    with get_webdriver() as driver:
        driver.get("https://servicos.efazenda.ms.gov.br/sgae/EmissaoDAEMSdeICMS/")
        time.sleep(5)  # Tempo para o certificado carregar (ajuste conforme necessário)
        
    # Navegação pelo formulário e preenchimento dos dados
    # Abaixo está o restante do fluxo do Selenium, adaptado para Streamlit.
    # Adicione uma espera para a janela de certificado carregar
        time.sleep(10)  # Ajuste conforme necessário para garantir que a caixa de diálogo apareça

    
        # Selecionar o botão de rádio com valor "Sim"
        radio_opcao_sim = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "Opcao"))
        )
        radio_opcao_sim.click()
        time.sleep(2)
        st.write("PASSEI POR AQUI")
        # Avançar para a próxima página
        botao_avancar = driver.find_element(By.ID, "avancar")
        botao_avancar.click()
        time.sleep(5)
        # Esperar que o campo select2 seja carregado e clique nele para abrir o menu
        select2_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="select2-chosen-1"]'))
        )
        select2_field.click()
    
        time.sleep(2)
        # Esperar que a barra de pesquisa apareça e inserir "310"
        search_bar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//input[contains(@class, "select2-input")]'))
        )
        search_bar.send_keys("310")
        time.sleep(1)  # Espera breve para garantir que o filtro seja aplicado
    
        # Pressionar Enter para selecionar a opção filtrada automaticamente
        search_bar.send_keys(Keys.ENTER)
        st.write("PASSEI POR AQUI")
        time.sleep(2)
    
        # Esperar que o botão "Avançar" esteja carregado e clique nele
        botao_avancar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="avancar"]'))
        )
        botao_avancar.click()
        time.sleep(2)
    
        # Esperar que o campo select2 seja carregado e clique nele para abrir o menu
        select2_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="select2-chosen-1"]'))
        )
        select2_field.click()
    
        # Esperar que a barra de pesquisa apareça e inserir "IE"
        search_bar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//input[contains(@class, "select2-input")]'))
        )
        search_bar.send_keys("IE")
        time.sleep(1)  # Breve espera para garantir que o filtro seja aplicado
        search_bar.send_keys(Keys.ENTER)
        st.write("PASSEI POR AQUI")
        # Localizar e preencher o campo "Número do Documento"
        numero_doc_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="NumeroDoc"]'))
        )
        numero_doc_field.click()
        time.sleep(1)
        numero_doc_field.send_keys(numero_doc_input)
        time.sleep(1)
        numero_doc_field.send_keys(Keys.ENTER)
        time.sleep(2)

    
        try:
            # Esperar que o botão de fechamento esteja visível e clicável e, então, clicar
            popup_close_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@data-dismiss="modal"]'))
            )
            popup_close_button.click()
            time.sleep(1)  # Espera breve para garantir que o popup foi fechado
        except Exception as e:
            print("Popup não encontrado ou já fechado:", e)
            
        # Localizar e preencher o campo "Valor"
        valor_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "Valor"))
        )
        valor_field.click()
        valor_field.clear()
        valor_field.send_keys(valor_input)
        
        pdfs_gerados = []  # Lista para armazenar os PDFs gerados
        st.write("PASSEI POR AQUI")
    for i, chave_nf_input in enumerate(chaves_nf, start=1):
        st.write(f"Processando chave de acesso ({i}/{len(chaves_nf)}): {chave_nf_input}")

        
        # Clicar no botão "Incluir"
        excluir_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "Excluir"))
        )
        # Criar uma instância de ActionChains
        actions = ActionChains(driver)

        # Executar o duplo clique
        actions.double_click(excluir_button).perform()
        time.sleep(2)
        
            # Verificar se o popup está presente na primeira iteração



        
        st.write(f"Processando chave de acesso: {chave_nf_input}")

        # Localizar e preencher o campo "Chave NF"
        chave_nf_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ChaveNF"))
        )
        chave_nf_field.click()
        chave_nf_field.send_keys(chave_nf_input)

        # Clicar no botão "Incluir"
        incluir_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "incluirGrid"))
        )
        incluir_button.click()
        time.sleep(2)

        # Setar o checkbox
        checkbox = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "cb_gridChaveNota"))
        )
        if not checkbox.is_selected():
            checkbox.click()  # Marca o checkbox se não estiver marcado
        time.sleep(2)

        # Clicar no botão "Avançar"
        avancar_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "consultar"))
        )
        avancar_button.click()
        time.sleep(8)
        st.write("PASSEI POR AQUI")
        # Nome do PDF final baseado no número da página
        numero_texto = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//text[@class="span8"]'))
        ).text

        # Extrair o número da página para usar no nome do arquivo
        match = re.search(r"N\. FISCAIS DE NS:0*(\d+)", numero_texto)
        if match:
            numero_pagina = int(match.group(1))
            nome_arquivo = f"ImprimirPdfDaems{numero_pagina}.pdf"
            st.write(f"Nome do arquivo gerado será: {nome_arquivo}")
        else:
            st.error("Número de página não encontrado.")
            
            return None

        # Clicar no botão "Imprimir DAEMS"
        try:
            imprimir_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//input[@value="Imprimir DAEMS"]'))
            )
            imprimir_button.click()
            time.sleep(8)  # Tempo para a página de visualização carregar completamente
        except Exception as e:
            st.error(f"Erro ao clicar no botão 'Imprimir DAEMS': {e}")
            
            return None

        # Caminho para o PDF baixado temporariamente
        original_pdf_path = os.path.join(temp_dir, "ImprimirPdfDaems.pdf")
        output_folder = "pdfs_emitidos"  # Pasta onde os PDFs serão armazenados

        # Verificar se o arquivo de origem existe
        if os.path.exists(original_pdf_path):
            try:
                # Definir o caminho final do PDF
                local_pdf_path = os.path.join(output_folder, nome_arquivo)
                
                # Remover o arquivo existente, se já houver
                if os.path.exists(local_pdf_path):
                    os.remove(local_pdf_path)

                # Copiar o arquivo baixado para o destino final
                shutil.copy2(original_pdf_path, local_pdf_path)
                st.write(f"PDF salvo em: {local_pdf_path}")

                # Remover o arquivo original após copiar
                os.remove(original_pdf_path)

                # Ler o PDF como BytesIO para disponibilizá-lo no Streamlit
                with open(local_pdf_path, "rb") as pdf_file:
                    pdf_bytes = BytesIO(pdf_file.read())
                pdf_bytes.seek(0)  # Retorna ao início do arquivo
                pdfs_gerados.append((nome_arquivo, pdf_bytes))  # Adiciona à lista de PDFs gerados
            except FileNotFoundError as e:
                st.error(f"Arquivo não encontrado: {e}")
            except PermissionError as e:
                st.error(f"Permissão negada ao acessar o arquivo: {e}")
            except Exception as e:
                st.error(f"Erro ao copiar ou processar o PDF: {e}")
        else:
            st.error(f"Erro: PDF não foi baixado. Arquivo esperado: {original_pdf_path}")
            
            return None

    
    time.sleep(4)
    return pdfs_gerados







def show_selenium_log(logpath: str):
    if os.path.exists(logpath):
        with open(logpath) as f:
            content = f.read()
        st.code(body=content, language='log', line_numbers=True)
    else:
        st.error('Arquivo de log do Selenium não encontrado.', icon='🔥')

    




# Interface Streamlit
st.title("Automação de Emissão de Guias")

# Entradas de dados para os campos do formulário
numero_doc_input = st.text_input("Digite o número do documento:", value="288689992")
valor_input = st.text_input("Digite o valor (ex: 2487,25):", value="0,00")
chaves_nf_input = st.text_area("Digite as chaves de acesso separadas por linha:", value="")
total_guias = st.number_input("Digite o número total de guias a serem emitidas:", min_value=1, value=3, step=1)

# Processar as chaves como uma lista
chaves_nf = chaves_nf_input.splitlines()

# Criar a pasta de saída
output_folder = "pdfs_emitidos"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Botão para iniciar a automação
if st.button("Iniciar Automação"):
    logpath = os.path.join(output_folder, "selenium.log")
    try:
        resultado = run_selenium(logpath, "https://servicos.efazenda.ms.gov.br/sgae/EmissaoDAEMSdeICMS/",
                                 numero_doc_input, valor_input, chaves_nf[:total_guias])
        if resultado:
            for nome_arquivo, pdf_bytes in resultado:
                pdf_path = os.path.join(output_folder, nome_arquivo)
                with open(pdf_path, "wb") as file:
                    file.write(pdf_bytes.getbuffer())
                st.write(f"PDF salvo: {pdf_path}")
        else:
            st.error("Nenhuma guia foi gerada.")
    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")

# Exibir PDFs gerados
st.write("Arquivos gerados:")
for pdf_file in os.listdir(output_folder):
    pdf_path = os.path.join(output_folder, pdf_file)
    with open(pdf_path, "rb") as file:
        st.download_button(label=f"Baixar {pdf_file}", data=file, file_name=pdf_file, mime="application/pdf")
# Botão para apagar todos os PDFs
if st.button("Apagar Todos os PDFs"):
    for pdf_file in os.listdir(output_folder):
        pdf_path = os.path.join(output_folder, pdf_file)
        os.remove(pdf_path)
    st.success("Todos os PDFs foram apagados.")


