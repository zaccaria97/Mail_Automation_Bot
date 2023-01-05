from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart

 # Import
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import smtplib
import datetime
import imaplib
import os
import csv
import time




#------------------------------------------------------------------------------
#                                   COSTANTI
#------------------------------------------------------------------------------

 # PULSANTE "NUOVO MESSAGGIO"
class_button_nuovo_messaggio = "src-components-button__btn--28s2a.sidebar-header__btn.src-components-button__accent--21lz3"
# ELEMENTO PIVOT DEL FORM. CLASSE PER GET
class_to = "react-tagsinput-input"
# NON UTILIZZATO, FOCUS TABBANDO DAL CAMPO "TO"
class_body_msg = "src-components-newRichEditor__editor--3uZCB"
# NON UTILIZZATO, FOCUS TABBANDO DAL CAMPO "TO"
class_obj = "src-mail-components-newMessage__objectInput--2G6nT"
# PULSANTE "INVIO". CLASSE PER GET
class_send = "src-components-button__btn--28s2a.src-mail-components-newMessage__sendBtn--B0ElS.src-components-button__accent--21lz3"
# URL DELLA POSTA IN ARRIVO
URL = 'https://webmail.monema.it/mail/u/INBOX'
# PULSANTE LOGIN
class_login = "src-components-button__btn--28s2a.src-login-login__btn--3tLt3.src-components-button__accent--21lz3"
#Toast di messaggio inviato correttamente
toast_Success = "src-components-toastCustom__toastTitle--1DS0M"
#------------------------------------------------------------------------------
#                                 PARAMETRI
#------------------------------------------------------------------------------
# Tempo di ritardo tra un invio e la compilazione della mail successiva
DELAY_DOPO_INVIO_SECONDI = 5
# Secondi tra la compilazione dei campi e l'invio effettivo della mail
DELAY_SCRITTURA_EMAIL_E_INVIO = 2

#------------------------------------------------------------------------------
#                                  METODI
#------------------------------------------------------------------------------

# Metodo che apre il browser e attende che venga effettuato il login, restituisce
# il driver aperto (Poi da passare alla funzione "invia_email")
def apri_browser_ed_effettua_login():
    # Apertura browser (Schermo intero per evitare cambi layout responsive)
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(URL)
    # Attesa
    time.sleep(5)

    #LOGIN AUTOMATIC0

    # 1. Scrittura campi login
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "." + "input-wrap__input")))
    ActionChains(driver).move_to_element(element).send_keys("...email...").send_keys(Keys.TAB).send_keys("...password...").perform();

    # 2. Scrittura invio
    time.sleep(DELAY_SCRITTURA_EMAIL_E_INVIO)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "." + class_login)))
    element.click()


    # Finchè sono nella pagina di login non faccio nulla
    print(driver.current_url)
    while driver.current_url != URL:
        True
        time.sleep(1)
    return driver

# Metodo che invia una mail
# 0. Apertura browser e login automatico
# 1. Click sul bottone "Nuovo messaggio per aprire la tendina"
# 2. Scrittura sul campo cc
# 3. Scrittura sul campo obj
# 4. Scrittura corpo email
# 5 Click sul bottone "Invia"
# 6 Attesa Toast di successo
def invia_email(oggetto, messaggio, to,driver):
    # 0. Apertura finestra alla massima dimensione (Evita cambi di layout responsive)
    driver.maximize_window()

    print(to)
    # 1. Click sul bottone "Nuovo messaggio"
    #driver.find_element(By.CLASS_NAME, class_button_nuovo_messaggio_).click()
    element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "." + class_button_nuovo_messaggio)))
    element.click()
    # 2. Scrittura campi
    element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "." + class_to)))
    ActionChains(driver).move_to_element(element).send_keys(to).send_keys(Keys.TAB).send_keys(Keys.TAB).send_keys(oggetto).send_keys(Keys.TAB).send_keys(messaggio).perform();
    # 5. Scrittura invio
    time.sleep(DELAY_SCRITTURA_EMAIL_E_INVIO)
    element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "." + class_send)))
    element.click()

    # 6. Verifico che la mail sia inviata correttamente
    element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "." + toast_Success)))

'''
##lettura credenziali OBSOLETO: LE CREDENZIALI SONO HARDCODED
with open('CredenzialiMittente.txt',"r") as f:
    credenziali = f.readlines()
if credenziali[0]=='' or credenziali[1]=='' or credenziali[2]=='':
    print('inserire i dati nel file delle credenziali!')
    exit
'''

if __name__=="__main__":
    #Sul file mail_counter è presente l'indice dell'ultima mail inviata.
    with open('mail_counter.txt',"r") as f:
        ultimoIndice = f.read()

    list_of_column_names=[]
    driver = apri_browser_ed_effettua_login()

    with open('Target.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter = ',')
    # loop to iterate through the rows of csv
        index=0
        for row in csv_reader:
            # adding the first row
            if index==0:
                list_of_column_names.append(row)
                index=index+1
            else:
                if row[0]!='' and row[1]!='' and row[2]!='':
                    if index>int(ultimoIndice)+1:
                        invia_email(row[0],row[1],row[2],driver)#,credenziali[0],credenziali[1],credenziali[2], credenziali[3],driver)
                        with open("mail_counter.txt","w") as file:
                            file.write(str(index))
                        time.sleep(DELAY_DOPO_INVIO_SECONDI) #Tempo di attesa per inviare la mail successiva
                index=index+1
        f.close()

    with open("mail_counter.txt","w") as file:
        file.write("-1")
