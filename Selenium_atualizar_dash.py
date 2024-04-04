#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
import time


# In[ ]:


servico = Service(ChromeDriverManager().install())


# In[ ]:


navegador = webdriver.Chrome(service=servico)
navegador.get('https://app.powerbi.com/groups/me/list?experience=power-bi')
time.sleep(5)


# In[ ]:


usuario = 'xxxxxxxxxxxxxxxx@dominio.com'
senha =  'xxxxxxxx'

#Efetuar login:
navegador.find_element('xpath','/html/body/div/div[2]/div[2]/div/div[1]/div[2]/input').send_keys(usuario)
navegador.find_element('xpath','/html/body/div/div[2]/div[2]/button').click()
time.sleep(5)


# In[ ]:


#Efetuar login com senha:
navegador.find_element('xpath','/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[3]/div/div[2]/input').send_keys(senha)
navegador.find_element('xpath','/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div/div[3]/div/div[2]/div/div[5]/div/div/div/div/input').click()
time.sleep(30)


# In[ ]:


navegador.find_element('xpath','/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-item-renderer-panel/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-item-renderer/tri-extension-page-outlet/div[2]/workspace-view/tri-workspace-view/trident-workspace-shell/mat-sidenav-container/mat-sidenav-content/mat-sidenav-container/mat-sidenav-content/workspace-list-view/tri-workspace-list-view/div/main/fluent-workspace/mat-sidenav-container/mat-sidenav-content/fluent-workspace-list/fluent-list-table-base/div/cdk-virtual-scroll-viewport/div[1]/div[16]/div[2]/div/span/a').click()
time.sleep(5)


# In[ ]:


contar = 0
while contar < 3000:
    try:
        print("Elemento encontrado.")
        navegador.find_element('xpath','/html/body/div[2]/div[4]/div/mat-dialog-container/div/div/error-dialog/mat-dialog-actions/section/button').click()
        time.sleep(3)
        navegador.find_element('xpath','/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-item-renderer-panel/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-item-renderer/tri-extension-page-outlet/div[2]/dataset-details-container/dataset-action-bar/action-bar/action-button[2]/button/span').click()
        time.sleep(3)
        navegador.find_element('xpath','/html/body/div[2]/div[4]/div/div/div/span[1]/button/span').click()
        time.sleep(15)
        contar = contar + 1
        print(contar)   
    except NoSuchElementException:
        navegador.find_element('xpath','/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-item-renderer-panel/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-item-renderer/tri-extension-page-outlet/div[2]/dataset-details-container/dataset-action-bar/action-bar/action-button[2]/button/span').click()
        time.sleep(3)
        navegador.find_element('xpath','/html/body/div[2]/div[4]/div/div/div/span[1]/button/span').click()
        time.sleep(15)
        contar = contar + 1
        print(contar)    
