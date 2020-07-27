#!/usr/bin/env python
# coding: utf-8

# In[24]:


from selenium import webdriver # Importar classe para inicializar o browser
from selenium.webdriver.common.action_chains import ActionChains # Importar a classe ActionChains responsável pelas manipulações
from selenium.webdriver.common.keys import Keys #Importar a classe Keys que podem ser utilizadas no key_up e key_down
from selenium.webdriver.support.ui import WebDriverWait # Importar a classe WebDriverWait
from selenium.webdriver.support import expected_conditions as EC # Importar a classe que contém as funções e aplicar um alias
from selenium.webdriver.common.by import By # Importar classe para ajudar a localizar os elementos
from datetime import datetime
import time
import os
import pyautogui
import pandas as pd
import ctypes #Criar tela de mensagem
import configparser
#pd.options.display.float_format = '{:,.2f}'.format
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning) 


# In[2]:


#switch_to_window --mudar a janela
#current_window_handle -- pega o handle da janela
#current_url -- verifica a url ativa do handle
#window_handles -- lista os handles abertos
#switch_to_window' -- muda a janela para o nandle
#driver.maximize_window()


# In[3]:


def AbreSitePrincipal():
    #inicia
    driver = webdriver.Chrome()
    driver.get('https://app.sefaz.go.gov.br/arr-www/view/entradaContribuinte.jsf')
    wait = WebDriverWait(driver, 1000)
    
    return driver


# In[4]:


def Inicial_Guia_icms_normal(driver):
    #Clicar no Botao icms - Normal
    driver.find_element_by_xpath('//*[@id="frmEntradaContribuinte:entradaContribuinte2:8:botaoEmissaoDARE"]').click()
    time.sleep(1)


# In[5]:


def Abre_inscricao_estadual(driver):

    # Mudar o handle
    driver.switch_to_window(driver.window_handles[-1]) #ativa a ultima aba aberta, para poder manipular seus elementos
    
    #Espera até que o elemento aparecça ate 10 segundos
    wait = WebDriverWait(driver, 1000)
    elemento = wait.until(EC.visibility_of_element_located((By.ID, 'frmEmissao:btnContinueBaixo')))
     #Seleciona contribuintes que tem Inscriçao estadual
    driver.find_element_by_xpath('//*[@id="frmEmissao:apContribuinte"]/h3[1]').click()
    time.sleep(2)


# In[6]:


def Preenche_inscricao_estadual(driver, sInscricaoEstadual):
    
    wait = WebDriverWait(driver, 1000) #Espera o site carregar em até 1000 segundos
    txtInscricao = driver.find_element_by_xpath('//*[@id="frmEmissao:apContribuinte:txtInscricao"]')
    
    txtInscricao.send_keys(Keys.HOME) #Posiciona no inicio
    txtInscricao.send_keys(sInscricaoEstadual)
    
    time.sleep(2) #espera 2 segundos antes de apertar o botão para continuar
    driver.find_element_by_id('frmEmissao:btnContinueBaixo').click()
    time.sleep(1) #espera 2 segundos antes de apertar o botão para continuar
        


# In[7]:


def Abre_Form_Dare(driver):
    #Espera até que o elemento apareca no maximo até 100 segundos
    wait = WebDriverWait(driver, 1000)
    elemento = wait.until(EC.visibility_of_element_located((By.ID, 'frmEmissao:apReceitaEstadual:treeReceita_node_0')))
    
    #forçar aparecer os elementos ICMS NORMAL -> 1-ICMS - IMPOSTO (...) -> 108 - Normal ->300 - Mensal
    for i in range(1,4):
        driver.execute_script(f'document.getElementsByTagName("ul")[{i}].style.display = "block"')
    time.sleep(0.5)
    
    #Escolhe 300 - Mensal
    driver.find_element_by_xpath('//*[@id="frmEmissao:apReceitaEstadual:treeReceita_node_0_0_0_5"]/div/span').click()
    time.sleep(1)
    
    #clicar em continuar
    driver.find_element_by_id("frmEmissao:btnContinueBaixo").click()
    time.sleep(1)


# In[8]:


def Preenche_Dare(driver, nAno, nMes, dtVenc, dtPgto,nValor, sComplemento):
    
    '''driver  => navegador criado pelo selenium
       nAno    => Ano como numero AAAA , será convertida nesta funcao
       nMes    => Mes como numero M, será convertido para texto nesta funcao
       dDtVenc => Data de vencimento. Será convertido para o formato correto AAAA/MM/DD nesta função
       dDtPgto => Data de pagameno. Será convertido para o formato correto AAAA/MM/DD nesta função
       nValor  => Valor do imposto. Será convertido para texto, nesta função
       sComplemento =>Complemento tem que ser texto, nenhum tratamento será dado
    '''
    sAno, sMes = str(nAno), str(nMes+1) # O mês será somado 1, pois na combo começa com 2
    sDtVenc, sDtPgto = dtVenc.strftime('%d/%m/%Y'), dtPgto.strftime('%d/%m/%Y')
    
    sValor = str(nValor).replace('.',',')
    
    wait = WebDriverWait(driver, 100)
    elemento = wait.until(EC.visibility_of_element_located((By.ID, 'frmEmissao:acInfoDocumento:cbMes')))   
    #Clica no seletor do mes
    driver.find_element_by_xpath('//*[@id="frmEmissao:acInfoDocumento:cbMes"]/input').click()
    time.sleep(0.5)
    #Escolhe o mês começa com o 2, por exemplo, janeiro e 2, /ul/li[2]
    driver.find_element_by_xpath('//*[@id="frmEmissao:acInfoDocumento:cbMes_panel"]/ul/li[' + sMes + ']').click()
    
    #Ano
    driver.find_element_by_xpath('//*[@id="frmEmissao:acInfoDocumento:txtAno"]/input').click()
    time.sleep(1)
    
    #pega a lista de anos
    lista_anos = driver.find_elements_by_xpath('//*[@id="frmEmissao:acInfoDocumento:txtAno_panel"]/ul/li')
    #Descobre o ano e clica nele, mas é importante que a lista esteja aparecendo
    for ano in lista_anos:
        if ano.text == sAno:
            ano.click()
    
    #Remove a trava do campo passando comando javascript
    driver.execute_script('document.getElementsByName("frmEmissao:acInfoDocumento:cdDataVencimento_input")[0].removeAttribute("readonly")')
    driver.execute_script('document.getElementsByName("frmEmissao:acInfoDocumento:cdDataPagamento_input")[0].removeAttribute("readonly")')

    driver.find_element_by_xpath('//*[@id="frmEmissao:acInfoDocumento:cdDataVencimento_input"]').send_keys(sDtVenc)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="frmEmissao:acInfoDocumento:cdDataPagamento_input"]').send_keys(sDtPgto)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="frmEmissao:acInfoDocumento:txtValorOriginal"]').send_keys(sValor)

    complemento = driver.find_element_by_xpath('//*[@id="frmEmissao:acInfoDocumento:txtInformacoesComplementares"]')
    #apaga complemento
    complemento.send_keys(Keys.CONTROL + Keys.SHIFT + Keys.HOME)#selecionar
    complemento.send_keys(Keys.BACKSPACE) #limpar
    #preenche complemento
    complemento.send_keys(sComplemento)

    #Gerar dare
    driver.find_element_by_xpath('//*[@id="frmEmissao:btnGerarBaixo"]').click()
    


# In[ ]:





# In[26]:


def Imprime(driver):
    
    time.sleep(5)#Esperar 5 segundos para dar tempo de carregar a visualizaçao do pdf, não estou conseguindo esperar os elementos serem carregados
    #IMPRESSAO
    driver.switch_to_window(driver.window_handles[-1])#ativar a tela da impressão
    driver.maximize_window()
    time.sleep(1)
    
    #Espera até que o elemento aparecça ate 10 segundos, ISTO NÃO FUNCIONOU
    #wait = WebDriverWait(driver, 20)
    #elemento = wait.until(EC.visibility_of_element_located((By.ID, 'toolbar')))
    
    #Ativar o menu pela imagem
    #posicao_aparecer_btn_imprimir = pyautogui.locateCenterOnScreen('imagens\\ForcarAparecerIconeImpressao.png',confidence=0.90)
    #pyautogui.moveTo(posicao_aparecer_btn_imprimir,duration=1)
    #posicao_imprimir_escondido    = pyautogui.locateCenterOnScreen('imagens\\btn_Imprimir_escondido.png',confidence=0.90)
    #pyautogui.click(posicao_imprimir_escondido,duration=1)
    
    #Ativa o menu de imprimr pleo menu
    tela_impressao = driver.find_element_by_xpath('/html/body')
    actions = ActionChains(driver)
    actions.context_click(tela_impressao)
    actions.send_keys(Keys.DOWN)
    actions.perform()
    
    #Ativa a impressão pelo menu
    posicao_menu_imprimir = pyautogui.locateCenterOnScreen('imagens\\menu_ImprimirControl-P.png',confidence=0.90)
    pyautogui.click(posicao_menu_imprimir,duration = 0.5)
  
    
    #Executa a impressão pelo botão
    time.sleep(3)
    posicao_btn_imprimir = pyautogui.locateCenterOnScreen('imagens\\btn_Imprimir.png',confidence=0.90)
    pyautogui.click(posicao_btn_imprimir,duration = 0.5)
    


# In[10]:


def cria_arquivo_pdf(driver, caminho_arq):
    #Se a impresao for  em PDF
    time.sleep(2)
    #nome_arq_full = 'dare' + '_' + datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    #caminho_arq = os.path.join(os.getcwd(),'guias','GO',nome_arq_full + '.pdf')
    
    #Nomeia o arquivo
    posicao_nome_arquivo = pyautogui.locateCenterOnScreen('imagens\\txt_nome_arquivo.png',confidence=0.90 )
    
    pyautogui.click(posicao_nome_arquivo,duration = 0.5)
    pyautogui.typewrite(caminho_arq, interval = 0.001)
    
    time.sleep(1)
    posicao_btn_salvar = pyautogui.locateCenterOnScreen('imagens\\btn_salvar.png',confidence=0.90 )
    pyautogui.click(posicao_btn_salvar,duration = 0.5) #salvar
    time.sleep(5) #5 segundos para dar tempo salvar
    


# In[11]:


def Reinicia(driver):
    
    driver.close() #Fecha a janela de impressao
    driver.switch_to_window(driver.window_handles[1])#ativa a Janela de preenchimento de inscriçao
    driver.close() #Fecha a Janela de preenchimento de inscriçao estadual
    driver.switch_to_window(driver.window_handles[0]) #Volta para a tela inicial
    return driver


# In[12]:


def df_dados():
    
    config = configparser.ConfigParser()
    config.read('config.ini')
    caminho_controle = config['path']['controlador_impostos']
    dtc = {'UF' : str, 'LOJA' : str, 'IE' : int, 'ANO BASE': int, 'MES BASE': int, 'VLR ICMS' : float, 'DT VENC' : str,'DT PAG' : str }
    dados = pd.read_excel(caminho_controle,sheet_name='GO',dtypes=dtc)
    
    dados['DT VENC'] = pd.to_datetime(dados['DT VENC'], format='%d/%m/%Y')
    dados['DT PAG'] = pd.to_datetime(dados['DT PAG'], format='%d/%m/%Y')
    
    return dados


# In[13]:


def Start():
    
    dados = df_dados()
    try:
        driver = AbreSitePrincipal()
        driver.maximize_window()
        for i in dados.index:
        #for i in range(3):
            sInscricaoEstadual = str(dados['IE'][i])
            nAno, nMes         = dados['ANO BASE'][i], dados['MES BASE'][i]
            dDtVenc, dDtPgto   = dados['DT VENC'][i] , dados['DT PAG'][i]
            nValor             = dados['VLR'][i]
            sComplemento       = dados['LOJA'][i]
            sCaminhoArquivo    = os.path.join(dados['DIRETORIO'][i], dados['ARQUIVO'][i] + '.pdf')
            
            if not (os.path.isfile(sCaminhoArquivo)): #Verifica se o arquivo existe
                Inicial_Guia_icms_normal(driver)
                Abre_inscricao_estadual(driver)
                Preenche_inscricao_estadual(driver,sInscricaoEstadual)
                Abre_Form_Dare(driver)
                Preenche_Dare(driver,nAno, nMes, dDtVenc, dDtPgto,nValor, sComplemento)
                Imprime(driver)
                cria_arquivo_pdf(driver,sCaminhoArquivo)
                driver = Reinicia(driver)
            #else:
            #    criar_log()
    except:
          pass
    finally:
        #finaliza
        time.sleep(5) #eperar 5 segundos por algum processo final
        driver.quit()  


# In[25]:


print("Inicio ... ========================")
Start()
print("Finalizado ! ======================")


# # Teste

# In[15]:


#driver = AbreSitePrincipal()


# In[16]:


#driver.maximize_window()


# In[17]:


#Inicial_Guia_icms_normal(driver)


# In[18]:


#Abre_inscricao_estadual(driver)


# In[19]:


#Preenche_inscricao_estadual(driver,"107832070")


# In[20]:


#Abre_Form_Dare(driver)


# In[21]:


#dados = df_dados()


# In[22]:


# i = 0
# sInscricaoEstadual = str(dados['IE'][i])
# nAno, nMes         = dados['ANO BASE'][i], dados['MES BASE'][i]
# dDtVenc, dDtPgto   = dados['DT VENC'][i] , dados['DT PAG'][i]
# nValor             = dados['VLR'][i]
# sComplemento       = dados['LOJA'][i]
# sCaminhoArquivo    = os.path.join(dados['DIRETORIO'][i], dados['ARQUIVO'][i] + '.pdf')


# In[23]:


#Preenche_Dare(driver,nAno, nMes, dDtVenc, dDtPgto,nValor, sComplemento)


# In[ ]:




