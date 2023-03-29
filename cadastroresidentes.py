import asyncio
import tkinter as tk
from tkinter import filedialog
import random
import os
import time
import pandas as pd
import undetected_chromedriver as uc
from anticaptchaofficial.recaptchav2proxyless import *
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

load_dotenv()
API_KEY=os.getenv('API_KEY')
CPF=os.getenv('CPF')
SENHA=os.getenv('SENHA')
Contador=0
tempo=round(random.uniform(1,4),2)

# Cria a janela principal
root = tk.Tk()
root.withdraw()

# Abre o diálogo para seleção de arquivo
file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

# Lê o arquivo Excel e carrega em um DataFrame pandas
df = pd.read_excel(file_path)

# Inicializa o navegador (Chrome, neste exemplo)
driver = uc.Chrome()

#driver = webdriver.Chrome()

# Abre o site
driver.get('http://sinar.mec.gov.br/')
try:
    # Aguarda até que o botão "Entrar" esteja disponível
    btn_entrar = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CLASS_NAME, 'btn.btn-block.btn-circle.btn-entrar'))
    )

    # Clica no botão "Entrar"
    btn_entrar.click()

    # Aguarda até que o botão para inserir CPF esteja disponível
    btn_cpf = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, 'accountId'))
    )

    # Clica no botão para inserir CPF
    btn_cpf.click()

    # Localiza o campo de CPF
    campo_cpf = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, 'accountId'))
    )
    time.sleep(tempo)
    # Insere o CPF no campo
    campo_cpf.send_keys(CPF)

    btn_entrar2 = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CLASS_NAME, 'button-continuar'))
    )
    btn_entrar2.click()

    # Localiza o campo de CPF
    campo_senha = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, 'password'))
    )

    time.sleep(tempo)
    # Insere o CPF no campo
    campo_senha.send_keys(SENHA)

    #Digita senha 
    btn_entrar3 = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, 'submit-button'))
    )
    btn_entrar3.click()

    #adicionar verificação de Captcha caso aconteça (estudar a recaptcha do GOV)


    #Localiza e espera o link de acesso a COREMU
    element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="/cnrms/index/14628"]'))
    )
    # Clica no elemento <a>
    element.click()

    # Aguarda até que o botão do menu esteja disponível
    btn_menu = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, 'menu-item-menu-8'))
    )
    # Clica no botão do menu
    btn_menu.click()

    # Aguarda até que o link para cadastrar residente esteja disponível
    link_cadastrar_residente = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, 'menu-funcionalidade-163'))
    )
    # Clica no link para cadastrar/pesquisar residente
    link_cadastrar_residente.click()

    #Clica no link para abrir página de cadastro
    cadastrar_residente = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="/residentes/gerenciar/inserir"]'))
    )
    cadastrar_residente.click()
    time.sleep(tempo)

    #Selecionando o frame mais a frente    
    driver.switch_to.frame(0)

    #Clicar no elemento do iframe do Recaptcha
    btn_captcha=WebDriverWait(driver,10).until(
    EC.presence_of_element_located((By.ID,'recaptcha-anchor-label'))
    )
    btn_captcha.click()
    time.sleep(tempo)

    # Volta para o frame pai
    driver.switch_to.parent_frame()
    driver.switch_to.default_content()  # opcional, para garantir que está no topo da estrutura de frames
    time.sleep(tempo)
    
    if len(driver.find_elements(By.XPATH,('//*[@id="rc-imageselect"]'))) > 0:
        sitekey = driver.find_element(By.XPATH, '//*[@id="formInserirResidente"]')
        sitekey_html = sitekey.get_attribute('outerHTML')
        sitekey_clean=sitekey_html.split('"><div style')[0].split('data-sitekey="')[1]
                
        #Define a URL do Recaptcha
        URL='http://sinar.mec.gov.br/residentes/gerenciar/inserir'
        solver = recaptchaV2Proxyless()
        solver.set_verbose(1)
        solver.set_key(API_KEY)
        solver.set_website_url(URL)
        solver.set_website_key(sitekey_clean)

        g_response = solver.solve_and_return_solution()
        if g_response!= 0:
            print("g_response"+g_response)
            #Faz o Navegador abrir a box de texto do Recaptcha
            driver.execute_script('var element=document.getElementById("g-recaptcha-response"); element.style.display="";')
                
            #adiciona o token do Recaptcha
            driver.execute_script("""document.getElementById("g-recaptcha-response").innerHTML = arguments[0]""", g_response)
                        
            #Oculta novamente a box de texto
            driver.execute_script('var element=document.getElementById("g-recaptcha-response"); element.style.display="none";')
            btn_verificar=WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="recaptcha-verify-button"]'))
            )
            btn_verificar.click()
            time.sleep (tempo)            
                    
            #Digita o CPF do residente na página de cadastro de residente            
            # cpf_residente=WebDriverWait(driver,10).until(
            #     EC.presence_of_element_located((By.ID, 'nuCpf'))
            #     )
            # cpf_residente.send_keys('05508882413')
            for index, row in df.iterrows():
                # Digita o CPF do residente
                cpf_residente = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'nuCpf'))
                )
                cpf_residente.send_keys(row['residenteMatriculado.cpf'])
    
                #Clica botão de pesquisar na página de cadastro de residente
                btn_cpf_residente=WebDriverWait(driver,10).until(
                EC.presence_of_element_located((By.ID,'btnPesquisar'))  
                )
                btn_cpf_residente.click()
                time.sleep(5)

                # Digita o e-mail do residente
                email_residente = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'dsEmail'))
                )
                email_residente.send_keys(row['residenteMatriculado.email'])
                time.sleep (tempo)
                    
                # Digita a data de nascimento do residente
                nascimento_residente = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'dtNascimento'))
                )
                nascimento_residente.send_keys(row['residente.dadosResidente.dadosPessoais.dataNascimento'])
                time.sleep (tempo)
                    
                # Digita a data de formatura do residente
                formatura_residente = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'dtFormatura'))
                )
                formatura_residente.send_keys(row['residente.dadosResidente.datadeformatura'])
                time.sleep (tempo)

                # Seleciona a opção correspondente à graduação do residente
                graduacao_residente = Select(driver.find_element(By.ID, 'noGraduacao'))
                graduacao_residente.select_by_visible_text(row['residenteMatriculado.categoriaProfissional.descricao'])
                time.sleep (tempo)

                # Digita o nome correspondente da instituição de formação
                residente_uf = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'noIes'))
                )
                residente_uf.send_keys(row['residente.dadosResidente.instituicao.nome'])
                time.sleep (tempo)
                        
                # Seleciona a opção correspondente à UF da Instituição
                residente_uf = Select(driver.find_element(By.ID, 'sgUf'))
                residente_uf.select_by_visible_text(row['residente.dadosResidente.instituição.uf'])
                time.sleep (tempo)
                    
                # Seleciona a opção correspondente ao tipo de formação
                residente_uf = Select(driver.find_element(By.ID, 'tpFormacao'))
                residente_uf.select_by_visible_text(row['residente.dadosResidente.tipo.formação'])
                time.sleep (tempo)

                # Seleciona a opção correspondente a nacionalidade da instituição
                residente_uf = Select(driver.find_element(By.ID, 'noPais'))
                residente_uf.select_by_visible_text(row['residente.dadosResidente.dadosPessoais.pais.descricao'])
                time.sleep (tempo)

                # Digita opção correspondente ao número de registro do Conselho do residente
                residente_uf = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'nuRegistroProfissional'))
                )
                residente_uf.send_keys(row['residente.dadosResidente.dadosProfissionais.numeroInscricao'])
                time.sleep (tempo)

                # Seleciona a opção correspondente à UF do Conselho do residente
                residente_uf = Select(driver.find_element(By.ID, 'sgUfRegistro'))
                residente_uf.select_by_visible_text(row['residente.dadosResidente.dadosProfissionais.numeroInscricao.uf'])
                time.sleep (tempo)

                # Seleciona a opção correspondente à área profissional do residente
                conselho_profissional_residente = Select(driver.find_element(By.ID, 'sgConselhoProfissional'))
                conselho_profissional_residente.select_by_visible_text(row['residenteMatriculado.categoriaProfissional.conselhoProfissional.descricao'])
                time.sleep (tempo)

                # Seleciona a opção correspondente à área profissional do residente
                area_profissional_residente = Select(driver.find_element(By.ID, 'coAreaProfissional'))
                area_profissional_residente.select_by_visible_text(row['residenteMatriculado.categoriaProfissional.descricao'])
                time.sleep (tempo)

                #clica no botão adicionar
                btn_addRegistro=WebDriverWait(driver,10).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="addRegistro"]'))
                    )
                btn_addRegistro.click()
                time.sleep (tempo)

                #clica no botão salvar
                btn_RegistroSalvar=WebDriverWait(driver,10).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="btnSalvar"]'))
                    )
                btn_RegistroSalvar.click()
                time.sleep (tempo)

                # Adiciona a informação do cadastro
                df.at[index, 'Resultado'] = 'Cadastrado'
                time.sleep(tempo)
        else:
            print("task finished with error"+solver.error_code)
                # Salva as alterações na planilha
        df.to_excel('residentes-atualizado', index=False)
          
    else:
        #Digita o CPF do residente na página de cadastro de residente            
        #cpf_residente=WebDriverWait(driver,10).until(
        #EC.presence_of_element_located((By.ID, 'nuCpf'))
        #)
        #cpf_residente.send_keys('05508882413')
            
        
        for index, row in df.iterrows():
            # Digita o CPF do residente
            cpf_residente = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'nuCpf'))
            )
            cpf_residente.send_keys(row['residenteMatriculado.cpf'])

            #Clica botão de pesquisar na página de cadastro de residente
            btn_cpf_residente=WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.ID,'btnPesquisar'))  
            )
            btn_cpf_residente.click()
            time.sleep(5)

            # Digita o e-mail do residente
            email_residente = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'dsEmail'))
            )
            email_residente.send_keys(row['residenteMatriculado.email'])
            time.sleep (tempo)
                
            # Digita a data de nascimento do residente
            nascimento_residente = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'dtNascimento'))
            )
            nascimento_residente.send_keys(row['residente.dadosResidente.dadosPessoais.dataNascimento'])
            time.sleep (tempo)
                
            # Digita a data de formatura do residente
            formatura_residente = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'dtFormatura'))
            )
            formatura_residente.send_keys(row['residente.dadosResidente.datadeformatura'])
            time.sleep (tempo)

            # Seleciona a opção correspondente à graduação do residente
            graduacao_residente = Select(driver.find_element(By.ID, 'noGraduacao'))
            graduacao_residente.select_by_visible_text(row['residenteMatriculado.categoriaProfissional.descricao'])
            time.sleep (tempo)

            # Digita o nome correspondente da instituição de formação
            residente_uf = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'noIes'))
            )
            residente_uf.send_keys(row['residente.dadosResidente.instituicao.nome'])
            time.sleep (tempo)
                    
            # Seleciona a opção correspondente à UF da Instituição
            residente_uf = Select(driver.find_element(By.ID, 'sgUf'))
            residente_uf.select_by_visible_text(row['residente.dadosResidente.instituição.uf'])
            time.sleep (tempo)
                
            # Seleciona a opção correspondente ao tipo de formação
            residente_uf = Select(driver.find_element(By.ID, 'tpFormacao'))
            residente_uf.select_by_visible_text(row['residente.dadosResidente.tipo.formação'])
            time.sleep (tempo)

            # Seleciona a opção correspondente a nacionalidade da instituição
            residente_uf = Select(driver.find_element(By.ID, 'noPais'))
            residente_uf.select_by_visible_text(row['residente.dadosResidente.dadosPessoais.pais.descricao'])
            time.sleep (tempo)

            # Digita opção correspondente ao número de registro do Conselho do residente
            residente_uf = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'nuRegistroProfissional'))
            )
            residente_uf.send_keys(row['residente.dadosResidente.dadosProfissionais.numeroInscricao'])
            time.sleep (tempo)

            # Seleciona a opção correspondente à UF do Conselho do residente
            residente_uf = Select(driver.find_element(By.ID, 'sgUfRegistro'))
            residente_uf.select_by_visible_text(row['residente.dadosResidente.dadosProfissionais.numeroInscricao.uf'])
            time.sleep (tempo)

            # Seleciona a opção correspondente à área profissional do residente
            conselho_profissional_residente = Select(driver.find_element(By.ID, 'sgConselhoProfissional'))
            conselho_profissional_residente.select_by_visible_text(row['residenteMatriculado.categoriaProfissional.conselhoProfissional.descricao'])
            time.sleep (tempo)

            # Seleciona a opção correspondente à área profissional do residente
            area_profissional_residente = Select(driver.find_element(By.ID, 'coAreaProfissional'))
            area_profissional_residente.select_by_visible_text(row['residenteMatriculado.categoriaProfissional.descricao'])
            time.sleep (tempo)

            #clica no botão adicionar
            btn_addRegistro=WebDriverWait(driver,10).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="addRegistro"]'))
                )
            btn_addRegistro.click()
            time.sleep (tempo)

            #clica no botão salvar
            btn_RegistroSalvar=WebDriverWait(driver,10).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="btnSalvar"]'))
                )
            btn_RegistroSalvar.click()
            time.sleep (tempo)

            # Adiciona a informação do cadastro
            df.at[index, 'Resultado'] = 'Cadastrado'
            time.sleep(tempo)
        df.to_excel('residentes-atualizado', index=False)

except Exception as e:
    print('Ocorreu um erro:', e)
    time.sleep(20)

finally:
    driver.quit()