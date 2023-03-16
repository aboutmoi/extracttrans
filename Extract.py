from cryptography import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import glob
import os
import shutil
import logging



logging.basicConfig(filename='C:/Users/PBBE05221/OneDrive - SNCF/Bureau/Analyse/Brouillon/Programme/V2/auto-py-to-exe-master/output/Historique_extraction.log', level=logging.INFO, format="%(asctime)s:%(levelname)s:%(message)s", encoding='utf-8')

logging.info("____________________Début programme____________________")

### Fonction qui stock dans une variable le fichier le plus récent du dossier téléchargement
def latest_download_file():
      path = r'C:/Users/PBBE05221/Downloads'
      os.chdir(path)
      files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
      newest = files[-1]
      return newest




### Récupération des données ITV Qlick ###
try:   
    driver = webdriver.Chrome(executable_path=r"C:/Users/PBBE05221/.wdm/drivers/chromedriver/win32/108.0.5359.71/chromedriver.exe")
    driver.get("https://digitop-qlik.sncf.fr/sense/app/370157a3-55d7-44d8-9179-f99f65515736/sheet/69873b99-bd04-4d32-8a27-ae577f152de6/state/analysis")
    driver.maximize_window()

    attente_until = WebDriverWait(driver, 80).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="grid"]/div[8]/div[1]/article/div[1]/div/div/div/div[1]/div[2]/table/tbody/tr[1]/th[1]/div/div/div/span'), 'N° intervention'))
    time.sleep(2)

    action = ActionChains(driver)

    Click_droit = driver.find_element(By.XPATH, '//*[@id="grid"]/div[8]/div[1]/article/div[1]/header')
    action.context_click(Click_droit)
    action.perform()

    attente_until = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="show-service-popup-dialog"]/div/div/div/div/ng-transclude/ul/li[3]')))

    Exporter = driver.find_element(By.XPATH, '//*[@id="show-service-popup-dialog"]/div/div/div/div/ng-transclude/ul/li[3]')
    action.click(Exporter)
    action.perform()

    attente_until = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="show-service-popup-dialog"]/div/div/div/div/ng-transclude/ul/li[4]')))

    Exporter_les_donnees = driver.find_element(By.XPATH, '//*[@id="show-service-popup-dialog"]/div/div/div/div/ng-transclude/ul/li[4]')
    action.click(Exporter_les_donnees)
    action.perform()

    attente_until = WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="export-dialog"]/div/div[2]/p[2]/a')))

    telechargement_csv = driver.find_element(By.XPATH, '//*[@id="export-dialog"]/div/div[2]/p[2]/a')
    action.click(telechargement_csv)
    action.perform()

    ### Vérifie que le fichier le plus récent est bien téléchargé
    fileends = "crdownload"
    while fileends == "crdownload":
        time.sleep(1)
        newest_file = latest_download_file()
        if "crdownload" in newest_file:
            fileends = "crdownload"
        else:
            fileends = "none"

    Fermer = driver.find_element(By.XPATH, '//*[@id="export-dialog"]/div/div[3]/button')
    action.click(Fermer)
    action.perform()

    print("! Fichier ITV Qlick téléchargé !")
    logging.info("Fichier ITV Qlick téléchargé !")


    ### Renommage du fichier téléchargé ###
    folder = os.listdir("C:/Users/PBBE05221/Downloads")
    newest = max(glob.iglob('C:/Users/PBBE05221/Downloads/*.xlsx'), key=os.path.getctime)
    src = newest
    dest = 'C:/Users/PBBE05221/Downloads/ITV.xlsx'
    os.rename(src, dest) 

    ### Déplacement du fichier IV ###
    filePath = shutil.copy("C:/Users/PBBE05221/Downloads/ITV.xlsx", "C:/Users/PBBE05221/SNCF/[ESTI IDF] GATI (Grp. O365) - Documents/Familles d'installations télécoms/Famille H Info Voyageurs Affichage/Dauphine/Analyse/")
    filePath = shutil.copy("C:/Users/PBBE05221/Downloads/ITV.xlsx", "C:/Users/PBBE05221/SNCF/[ESTI IDF] GATI (Grp. O365) - Documents/Familles d'installations télécoms/Famille N EAS Vidéo CADI/Vidéo Protection/Analyse/")

    ### Suppression du fichier IV dans son dossier d'origine ###
    os.remove('C:/Users/PBBE05221/Downloads/ITV.xlsx')

    print("! Fichier ITV Qlick renommé et déplacé avec succès !") 
    logging.info("Fichier ITV Qlick renommé et déplacé avec succès !")

except:
    print("Programme abrogé au niveau de l'extraction des ITV")
    logging.info("Programme abrogé au niveau de l'extraction des ITV")
    driver.quit()
    sys.exit()






### Récupération des données DI Qlick ###
try:   
    driver.get("https://digitop-qlik.sncf.fr/sense/app/4703c3d6-d136-4125-a251-3bdc34f23e18/sheet/1ba16cb4-936f-480f-9dad-19ae5985f3f3/state/analysis")

    attente_until = WebDriverWait(driver, 60).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="grid"]/div[8]/div[1]/article/div[1]/div/div/div/div[1]/div[2]/table/tbody/tr/th[1]/div/div/div/span'), 'N°DI'))
    time.sleep(2)

    Click_droit = driver.find_element(By.XPATH, '//*[@id="grid"]/div[8]/div[1]/article/div[1]/header')
    action.context_click(Click_droit)
    action.perform()

    attente_until = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="show-service-popup-dialog"]/div/div/div/div/ng-transclude/ul/li[3]')))

    Exporter = driver.find_element(By.XPATH, '//*[@id="show-service-popup-dialog"]/div/div/div/div/ng-transclude/ul/li[3]')
    action.click(Exporter)
    action.perform()

    attente_until = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="show-service-popup-dialog"]/div/div/div/div/ng-transclude/ul/li[4]')))

    Exporter_les_donnees = driver.find_element(By.XPATH, '//*[@id="show-service-popup-dialog"]/div/div/div/div/ng-transclude/ul/li[4]')
    action.click(Exporter_les_donnees)
    action.perform()

    attente_until = WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="export-dialog"]/div/div[2]/p[2]/a')))

    telechargement_csv = driver.find_element(By.XPATH, '//*[@id="export-dialog"]/div/div[2]/p[2]/a')
    action.click(telechargement_csv)
    action.perform()


    fileends = "crdownload"
    while fileends == "crdownload":
        time.sleep(1)
        newest_file = latest_download_file()
        if "crdownload" in newest_file:
            fileends = "crdownload"
        else:
            fileends = "none"


    Fermer = driver.find_element(By.XPATH, '//*[@id="export-dialog"]/div/div[3]/button')
    action.click(Fermer)
    action.perform()

    print("! Fichier DI Qlick téléchargé !")
    logging.info("Fichier DI Qlick téléchargé !")


    ### Renommage du fichier DI téléchargé ###
    folder = os.listdir("C:/Users/PBBE05221/Downloads")
    newest = max(glob.iglob('C:/Users/PBBE05221/Downloads/*.xlsx'), key=os.path.getctime)
    src = newest
    dest = 'C:/Users/PBBE05221/Downloads/DI.xlsx'
    os.rename(src, dest) 

    ### Déplacement du fichier DI ###
    filePath = shutil.copy("C:/Users/PBBE05221/Downloads/DI.xlsx", "C:/Users/PBBE05221/SNCF/[ESTI IDF] GATI (Grp. O365) - Documents/Familles d'installations télécoms/Famille H Info Voyageurs Affichage/Dauphine/Analyse/")
    filePath = shutil.copy("C:/Users/PBBE05221/Downloads/DI.xlsx", "C:/Users/PBBE05221/SNCF/[ESTI IDF] GATI (Grp. O365) - Documents/Familles d'installations télécoms/Famille N EAS Vidéo CADI/Vidéo Protection/Analyse/")

    ### Suppression du fichier DI dans son dossier d'origine ###
    os.remove('C:/Users/PBBE05221/Downloads/DI.xlsx')

    print("! Fichier DI Qlick renommé et déplacé avec succès !") 
    logging.info("Fichier DI Qlick renommé et déplacé avec succès !")

except:
    print("Programme abrogé au niveau de l'extraction des DI")
    logging.info("Programme abrogé au niveau de l'extraction des DI")
    driver.quit()
    sys.exit()






### Récupération des données WEBRMA ###
try:    
    driver.get("https://secteur-logistique-ti.mn.sncf.fr/controller/rma/visuFiches_refactored.php")
    driver.refresh()

    attente_until = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, 'acceptRGPD')))

    action2 = ActionChains(driver)

    Accept_RGPD = driver.find_element(By.ID, 'acceptRGPD')
    action2.click(Accept_RGPD)
    action2.perform()

    #Les 4 lignes suivantes sont à supprimer lorsque l'alerte de perturbation au SC ne sera plus afficher
    attente_until = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, 'buttonMoldalAlert')))
    
    Accept_Info_perturbation = driver.find_element(By.ID, 'buttonMoldalAlert')
    action2.click(Accept_Info_perturbation)
    action2.perform()

    attente_until = WebDriverWait(driver, 120).until(EC.text_to_be_present_in_element((By.ID, 'tableRmaAjax_info'), 'Enregistrement'))

    Extract_RMA = driver.find_element(By.ID, 'btnExport')
    action2.click(Extract_RMA)
    action2.perform()

    while "crdownload" not in latest_download_file() :
        time.sleep(1)

    fileends = "crdownload"
    while fileends == "crdownload":
        time.sleep(1)
        newest_file = latest_download_file()
        if "crdownload" in newest_file:
            fileends = "crdownload"
        else:
            fileends = "none"

    driver.quit()

    print("! Fichier RMA téléchargé !") 
    logging.info("Fichier RMA téléchargé !")


    ### Renommage du fichier téléchargé ###
    folder = os.listdir("C:/Users/PBBE05221/Downloads")
    newest = max(glob.iglob('C:/Users/PBBE05221/Downloads/*.csv'), key=os.path.getctime)
    src = newest  
    dest = 'C:/Users/PBBE05221/Downloads/RMA.csv'
    os.rename(src, dest)


    ### Déplacement du fichier RMA ###
    filePath = shutil.copy("C:/Users/PBBE05221/Downloads/RMA.csv", "C:/Users/PBBE05221/SNCF/[ESTI IDF] GATI (Grp. O365) - Documents/Familles d'installations télécoms/Famille H Info Voyageurs Affichage/Dauphine/Analyse/")
    filePath = shutil.copy("C:/Users/PBBE05221/Downloads/RMA.csv", "C:/Users/PBBE05221/SNCF/[ESTI IDF] GATI (Grp. O365) - Documents/Familles d'installations télécoms/Famille N EAS Vidéo CADI/Vidéo Protection/Analyse/")


    ### Suppression du fichier dans son dossier d'origine ###
    os.remove('C:/Users/PBBE05221/Downloads/RMA.csv')

    print("! Fichier RMA renommé et déplacé avec succès !")
    logging.info("Fichier RMA renommé et déplacé avec succès !") 

except:
    print("Programme abrogé au niveau de l'extraction des Fiches RMA")
    logging.info("Programme abrogé au niveau de l'extraction des Fiches RMA")
    driver.quit()
    sys.exit()

logging.info("____________________Fin programme____________________")

