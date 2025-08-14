import os
import requests
import yaml
import re
from datetime import datetime
from qreader import QReader
from cv2 import imreadmulti, imwritemulti, resize, INTER_AREA,IMWRITE_TIFF_COMPRESSION, IMREAD_GRAYSCALE, IMREAD_COLOR

from helpers import check_folders, get_unique_list_filename, move_file, set_logging, get_graph_access_headers, get_sharepoint_site_id, get_sharepoint_library_ids, get_list_columns, get_columns_mapping, normalize_qr_data, delete_file

    # Nahrání konfiguračního souboru
os.chdir(os.getcwd())
print(os.getcwd())
with open("config.yaml", "r", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

    #Kontrola a vytvoření složek
check_folders(cfg)

    #Iniciace logování
log = set_logging(cfg)

    # Hlavní program
log.info("Start")
start_time = datetime.now()

    # Získání přístupového tokenu M365 Graph, nastavení Bearer hlaviček
graph_headers = get_graph_access_headers(cfg)

    # Získání Sharepoint Site ID
site_id = get_sharepoint_site_id(cfg, graph_headers)

    # Získání Sharepoint library drive ID a list ID
drive_id, list_id = get_sharepoint_library_ids(cfg, site_id, graph_headers)

    # Nastavení Sharepoint drive URL pro upload souborů
drive_url = f"{cfg['graph_api']['drives']}{drive_id}/items/root:/"

    # Získání názvů sloupců library / listu ze Sharepoint
list_columns_names = get_list_columns(cfg, site_id, list_id, graph_headers)
log.debug("Názvy sloupců listu Sharepointu: {0}".format(list_columns_names))

    # Získání slovníku - mapy názvu sloupců name:displayName listu, ověření shody displayName z QR (konfigurační soubor) a listu Sharepointu  
columns_names_map_dict = get_columns_mapping(cfg, list_columns_names)
log.debug("Slovník mapování názvů sloupců a polí QR dat: {0}".format(columns_names_map_dict))

    # Inicializace QReaderu - rozpoznávání QR kódů
qr_reader = QReader()

    # Vytvoření seznamu souborů skenů ke zpracování
files_for_processing = [
    os.path.join(cfg['app']['source_path'], file_name)
    for file_name in os.listdir(cfg['app']['source_path'])
    if os.path.isfile(os.path.join(cfg['app']['source_path'], file_name))
    and os.path.splitext(file_name)[1] in cfg['app']['allowed_extensions']
]

    # Vyčištění složky "upload_path" - odstranění všech souborů
upload_files = os.listdir(cfg['app']['upload_path'])
if len(upload_files) > 0:
    log.info("Vyprazdňuji složku {0}".format(cfg['app']['upload_path']))
    for file_name in upload_files:
        file_path = os.path.join(cfg['app']['upload_path'], file_name)
        if os.path.isfile(file_path):
            os.remove(file_path)
            log.info("Smazán soubor {0}".format(file_path))  

    # Zpracování skenů, příprava k uploadu do Sharepointu
for file in files_for_processing:
    log.info("Zpracovávám {0}".format(file))
        # Načtení skenu ze soboru
    images = []
    ret, images = imreadmulti(mats=images, filename = file, flags=IMREAD_GRAYSCALE if cfg['app']['convert_to_grayscale'] else IMREAD_COLOR)
        # Kontrola načtení skenu, v případě chyby přesunout soubor do "unprocessed_path" a pokračovat dalším
    if not ret:
        moved_file = move_file(cfg['app']['unprocessed_path'], file)
        log.error("Nepodařilo se načíst soubor {0}. Nezpracován, přesunut do {1}.".format(file, moved_file))
        continue
    pages_count = len(images)
    log.info("Načteno stran: {0}".format(pages_count))

        # Načtení QR kód z první strany skenu
    qr_data = qr_reader.detect_and_decode(image=images[0])
    
        # Pokud QR kód neexistuje, přesunout soubor do "unprocessed_path" a pokračovat dalším
    if qr_data == ():
        moved_file = move_file(cfg['app']['unprocessed_path'], file)
        log.error("V souboru {0} nerozpoznán QR kód. Nezpracován, přesunut do {1}".format(file, moved_file))
        continue
    
        # Pokud první QR není validní (neobsahuje |), přesunout soubor do "unprocessed_path" a pokračovat dalším
    if qr_data[0].find("|") == -1: 
        moved_file = move_file(cfg['app']['unprocessed_path'], file)
        log.error("V QR kódu souboru {0} nerozpoznána validní data. Nezpracován, přesunut do {1}.".format(file, moved_file))
        continue            
            
    qr_data_list = []  # Inicializace prázdného listu pro QR data
    qr_data_list = qr_data[0].split("|") # Převod QR dat na list, oddělovač "|"
    log.debug("Načteno {0} polí z QR kódu: {1}".format(len(qr_data_list),qr_data_list))

        # Normalizace QR dat
    qr_data_formated, dif = normalize_qr_data(cfg, qr_data_list)
    log.debug('Počet polí QR dat po normalizaci: {0}'.format(len(qr_data_formated)))
    log.debug('Normalizovaná QR data: {0}'.format(qr_data_formated))

        # Změna velikosti skenů, je-li v konfiguračním souboru resize_scale různo od 1
    if cfg['app']['resize_scale'] != 1:
        log.debug("Zmenšuji velikost obrazů na {0} %".format(cfg['app']['resize_scale'] * 100))
        small_images = []
        for img in images:
            scale = cfg['app']['resize_scale']
            img_small = resize(img, None, fx=scale, fy=scale, interpolation=INTER_AREA)
            small_images.append(img_small)

        # Unikátní název souboru pro upload do Sharepointu - pole 3 QR kódu "Číslo výrobní zakázky"
    upload_file_name = get_unique_list_filename(f"{qr_data_list[2]}.tif", drive_url, graph_headers)
    log.debug("Unikátní název souboru pro upload: {0}".format(upload_file_name))

        # Vytvoření souboru pro upload a uložení do "upload_path" 
    upload_file_name_path = os.path.join(cfg['app']['upload_path'], upload_file_name)
    res = imwritemulti(upload_file_name_path, small_images, params=(IMWRITE_TIFF_COMPRESSION, 5))
    if res:
        log.info("Vytvořen soubor pro upload {0}, počet stran {1}.".format(upload_file_name_path, len(small_images)))
    else:
        log.warning("Nepodařilo se vytvořit soubor {0} pro upload, pokračuji dalším souborem.".format(upload_file_name_path))
        continue    

        # Nahrání souboru do Sharepointu
    with open(upload_file_name_path, "rb") as f:
        upload_url = f"{drive_url}{upload_file_name}:/content"
        response = requests.put(upload_url, headers=graph_headers, data=f)

            # Kontrola úspěšného nahrání souboru do Sharepointu. Při chybě přesun souboru do "unprocessed_path" a pokračovat dalším
        if response.status_code != 201:  
            moved_file = move_file(cfg['app']['unprocessed_path'], upload_file_name_path)
            log.error("{0} se nepodařilo nahrát do Sharepointu, přesunut do {1}.".format(upload_file_name_path, moved_file))
            continue

            # Získání id nahraného souboru
        etag = response.json()["eTag"].lower()
        file_uuid = re.findall("[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", etag)[0]
        file_url = f"{cfg['graph_api']['drives']}{drive_id}/items/{file_uuid}/listItem?select=id"
        response = requests.get(file_url, headers=graph_headers)
        
            # Kontrola úspěšného získání id nahraného souboru, při neúspěchu nelze nahrát metadata a pokračuje se dalším souborem
        if response.status_code != 200:
            log.error("Nepodařilo se získat Sharepoint id nahraného souboru {0}, nelze uložit metadata. Pokračuji dalším souborem)".format(upload_file_name))
            continue
        else:
            log.debug("Získáno id {1} nahraného souboru {0}".format(upload_file_name, response.json()["id"]))
            file_id = response.json()["id"]
            metadata_url = f"{cfg['graph_api']['sites']}{site_id}/lists/{list_id}/items/{file_id}/fields" # URL metadat (fields) nahraného souboru

            # Nahrání metadat (fields) k uloženému souboru z qr_data_formated
        results = []
        idx = 0
        for column_name in columns_names_map_dict.keys():
            column_value = {}
            column_value[column_name] = qr_data_formated[idx]
        
            # Nahrání metadat do sloupce Sharepointu, nejsou-li prázdná
            if qr_data_formated[idx] != "":
                log.debug("Nahrávám metadate do sloupce {0}".format(columns_names_map_dict[column_name]))
                metadata_response = requests.patch(metadata_url, headers=graph_headers, json=column_value)
                results.append(metadata_response.status_code)      
            else:
                log.debug("Sloupec {0} je prázdný, přeskočeno nahrání metadat.".format(columns_names_map_dict[column_name]))
            idx += 1
    
        results_ok_number = results.count(200)

        if results_ok_number == len(columns_names_map_dict):
            log.info("Všechna metadata {0} nahrána".format(upload_file_name))
        else:
         log.warning("Nahráno pouze {0} z {1} metadat {2} ".format(results_ok_number, len(columns_names_map_dict), upload_file_name))

        # Smazání originálního i zmenšeného souboru po úspěšném nahrání
    log.info("{0} {1}".format(file, "smazán" if delete_file(file) else "neexistuje"))
    log.info("{0} {1}".format(upload_file_name_path, "smazán" if delete_file(upload_file_name_path) else "neexistuje"))
        
duration = datetime.now() - start_time
log.info("Běh programu {0} sec".format(duration.seconds))