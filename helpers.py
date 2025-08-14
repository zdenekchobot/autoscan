from datetime import datetime
import re
import os
import logging
import requests

    # Získání unikátního názvu souboru - pokud již existuje, přidá se číselný index (x)
def get_unique_filename(file_name):
    idx = 0
    
    while os.path.exists(file_name):
        f_name,f_ext = os.path.splitext(file_name)   
        if re.findall("\(\d+\)$", f_name):
            idx = int(re.findall("\((\d+)\)$", f_name)[0]) + 1
            f_name = re.sub("\(\d+\)$", "", f_name)

        file_name = "{0}({1}){2}".format(f_name, str(idx), f_ext)
        idx += 1
    return file_name

    # Smazání souboru
def delete_file(file_name):
    if os.path.exists(file_name):
        os.remove(file_name)
        return True
    else:
        return False

    # Kontrola existence názvu souboru v Sharepointu 
def check_list_filename(file_name, drive_url, headers):
    name_url = drive_url + file_name + "?$select=name"
    response = requests.get(name_url, headers=headers)
    return False if response.status_code == 404 else True

    # Získání unikátního názvu souboru v Sharepointu
def get_unique_list_filename(file_name, drive_url, headers):
    idx = 0
    f_name, f_ext = os.path.splitext(file_name)
    while check_list_filename(file_name, drive_url, headers):
        file_name = "{0}({1}){2}".format(f_name, str(idx), f_ext)
        idx += 1
    return file_name
    
    # Kontrola a vytvoření složek
def check_folders(cfg):
    log_file_path, log_file_name = os.path.split(cfg['app']['log_file'])
    for path in [cfg['app']['unprocessed_path'], cfg['app']['upload_path'], log_file_path]:
        if not os.path.exists(path):
            os.makedirs(path)

    # Nastavení logování
def set_logging(cfg):
    log_format = logging.Formatter("%(asctime)s [%(levelname)-5.5s] %(message)s")
    log = logging.getLogger()
    log.setLevel(level=cfg['app']['log_level'].upper())
    file_log = logging.FileHandler(cfg['app']['log_file'])
    file_log.setFormatter(log_format)
    log.addHandler(file_log)
    console_log = logging.StreamHandler()
    console_log.setFormatter(log_format)
    log.addHandler(console_log)
    return log

    # Získání přístupového tokenu M365 Graph
def get_graph_access_headers(cfg):
    auth_url = f"https://login.microsoftonline.com/{cfg['auth']['tenant_id']}/oauth2/v2.0/token"
    auth_data = {
        "grant_type": "client_credentials",
        "client_id": cfg['auth']['client_id'],
        "client_secret": cfg['auth']['client_secret'],
        "scope": cfg['auth']['scope'],
    }
    auth_response = requests.post(auth_url, data=auth_data)
    access_token = auth_response.json().get("access_token")
    if not access_token:
        raise Exception("Nepodařilo se získat přístupový token MS Graph.")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    return headers  

    # Získání Sharepoint Site ID
def get_sharepoint_site_id(cfg, headers):
    site_url = f"{cfg['graph_api']['sites']}{cfg['m365']['hostname']}:/sites/{cfg['m365']['site_name']}?select=sharepointIds"
    response = requests.get(site_url, headers=headers)
    if response.status_code == 200:  # Kontrola získání Site ID
        return response.json()["sharepointIds"]["siteId"]
    else:
        raise Exception("Nepodařilo se získat Site ID")

    # Získání Sharepoint library drive ID a list ID
def get_sharepoint_library_ids(cfg, site_id, headers):
    library_url = f"{cfg['graph_api']['sites']}{site_id}/drives?$expand=list&search=name:{cfg['m365']['library_name']}"
    response = requests.get(library_url, headers=headers)
    if response.status_code == 200:  # Kontrola získání Library ID
        drive_id = response.json()["value"][0]["id"]
        list_id = response.json()["value"][0]["list"]["id"]
        return drive_id, list_id
    else:
        raise Exception("Nepodařilo se získat library drive nebo list ID")
    
    # Získání názvù sloupců knihovny/listu
def get_list_columns(cfg, site_id, list_id, headers):
    library_columns_url = f"{cfg['graph_api']['sites']}{site_id}/lists/{list_id}/columns?$select=displayName,name"
    response = requests.get(library_columns_url, headers=headers)
    if response.status_code == 200:  # Kontrola získání Library ID
        return response.json()["value"]
    else:
        raise Exception("Nepodařilo se získat seznam sloupců knihovny / listu")

    # Získání mapy názvu sloupců name:displayName
def get_columns_mapping(cfg, list_columns_names):
        # Kontrola konfiguračního souboru
    if not cfg['m365']['column_names']:
        raise Exception("V konfiguračním souboru chybí názvy sloupců pro mapování.")
    
        # Vytvoření slovníku pro mapování názvů sloupců
    list_columns_names_dict = {item["displayName"]: item["name"] for item in list_columns_names}
    
        # Vytvoření prázdného slovníku - mapy displayName:name sloupců knihovny
    columns_map_dict = {}
    
        # Ověření existence názvu sloupce Konfigurační soubor vs. Sharepoint list, naplnění mapy   
    for column_name in cfg['m365']['column_names']:
        if column_name in list_columns_names_dict:
            columns_map_dict[list_columns_names_dict[column_name]] = column_name
        else:
            raise Exception("Nepodařilo se namapovat sloupec {0} na sloupec v Sharepointu.".format(column_name))
    return columns_map_dict

    # Přesunutí souboru do "destination_path" s přejmenováním souboru na unikátní jméno, pokud již existuje
def move_file(destination_path, file_name):
    target_file_name = os.path.join(destination_path, os.path.basename(file_name))
    
        # Kontrola, zda soubor již existuje v cílové složce. Pokud ano, změní target_file_name na unikátní jméno
    if os.path.exists(target_file_name):
        target_file_name = get_unique_filename(target_file_name)
    
    os.rename(file_name, target_file_name)
    return target_file_name

    # Normalizace QR dat
def normalize_qr_data(cfg, qr_data_list):   
        # Vymazání nežádoucích QR dat
    qr_data_list.remove("$")
    
        # Normalizace délky listu doplněním polí s mezerami podle počtu sloupců v konfiguračním souboru listu Sharepointu
    dif = len(cfg['m365']['column_names']) - len(qr_data_list)
    if dif > 0:
        qr_data_list.extend([" "] * dif)

        # Náhrada prázdných hodnot mezerami
    qr_data_list = [item if item != "" else " " for item in qr_data_list]

    # Převod datumu "Uvažovaný termín balení" na ISO formát
    data_idx = 5
    if (len(re.findall(r"[0|1|2|3]?[0-9]\.[0|1]?[0-9]\.[1|2][0-9]{3}", qr_data_list[data_idx])) >= 0): # Je hodnota datum ve formátu dd.mm.yyyy?
        iso_date = datetime.strptime(qr_data_list[data_idx], "%d.%m.%Y").isoformat()
        qr_data_list[data_idx] = iso_date
    return qr_data_list, dif
 