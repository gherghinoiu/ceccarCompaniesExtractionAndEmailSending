import requests
import pandas as pd
import os
import sys
import time
import json 

# --- Setări Globale ---

TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_files')
API_URL = 'https://raportare.ceccar.ro/api/search'
COLOANE_DORITE = ['email', 'name', 'cui', 'region', 'phone', 'type']
HEADERS = {
    'Accept': 'application/ld+json',
    'Content-Type': 'application/ld+json',
    'Origin': 'https://raportare.ceccar.ro',
    'Referer': 'https://raportare.ceccar.ro/search',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Google Chrome";v="141", "Not?A_Brand";v="8", "Chromium";v="141"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"'
}

# --- FUNCȚIA 1: Extrage TOATE paginile (Cu Calea Corectă) ---

def extract_all_pages_for_region(region_id):
    output_filename = f'ceccar_data_regiunea_{region_id}_FULL.xlsx'
    print(f"Încep extragerea completă pentru regiunea ID: {region_id}...")
    
    all_items = []
    page = 1
    total_pages = 1
    
    try:
        while page <= total_pages:
            payload = {
                "page": page,
                "membersType": "companies",
                "memberLastName": "",
                "memberFirstName": "",
                "memberRegNumber": "",
                "memberRegion": region_id,
                "memberCurrentYearVisa": None
            }
            
            if total_pages > 1:
                print(f"Se procesează pagina {page} din {total_pages}...")
            else:
                print("Se procesează prima pagină...")

            response = requests.post(API_URL, headers=HEADERS, json=payload)
            response.raise_for_status()
            
            # Folosim .json() direct, așa cum ar fi trebuit de la început
            data = response.json()
            
            # --- SOLUȚIA CORECTĂ ---
            # Căutăm 'items' ÎN INTERIORUL 'pager'
            items = data.get('pager', {}).get('items', [])
            # --- SFÂRȘIT SOLUȚIE ---
            
            if page == 1:
                # Calea pentru paginare este DE ASEMENEA în interiorul 'pager'
                pagination_data = data.get('pager', {}).get('pagination', {})
                total_pages = pagination_data.get('total_pages', 1)
                
                print(f"Detectat un total de {total_pages} pagini.")
                if not items and total_pages > 0:
                    print(f"Eroare: Serverul a raportat pagini, dar lista 'items' este goală.")
                    return
                elif not items:
                    print("Nu s-au găsit date pentru această regiune.")
                    return
            
            if items:
                all_items.extend(items)
            
            page += 1
            if page <= total_pages:
                time.sleep(0.25)

        print(f"\nExtragere completă finalizată. Total item-uri colectate: {len(all_items)}")
        
        df = pd.DataFrame(all_items)
        for col in COLOANE_DORITE:
            if col not in df.columns:
                df[col] = None
        df = df[COLOANE_DORITE]

        df.to_excel(output_filename, index=False, sheet_name='CECCAR Data')
        print(f"Datele au fost salvate cu succes în fișierul: {output_filename}")

    except Exception as e:
        print(f"A apărut o eroare în `extract_all_pages_for_region`: {e}")


# --- FUNCȚIA 2: Extrage o SINGURĂ pagină (Cu Calea Corectă) ---

def extract_single_page(region_id, page_number, verbose=False):
    """
    Extrage datele pentru o singură pagină, folosind calea JSON corectă.
    """
    print(f"Încep extragerea pentru Regiunea ID: {region_id}, Pagina: {page_number}...")
    
    if not os.path.exists(TEMP_DIR):
        os.makedirs(TEMP_DIR)
        print(f"Creat folderul: {TEMP_DIR}")
    
    payload = {
        "page": page_number,
        "membersType": "companies",
        "memberLastName": "",
        "memberFirstName": "",
        "memberRegNumber": "",
        "memberRegion": region_id,
        "memberCurrentYearVisa": None
    }
    
    if verbose:
        print("\n" + "="*50)
        print("--- DETALII APEL API (TRIMITERE) ---")
        print(f"URL: {API_URL}")
        print(f"HEADERS: \n{json.dumps(HEADERS, indent=2)}")
        print(f"PAYLOAD: \n{json.dumps(payload, indent=2)}")
        print("="*50 + "\n")
        
    try:
        response = requests.post(API_URL, headers=HEADERS, json=payload)
        
        if verbose:
            print("\n" + "="*50)
            print("--- DETALII RĂSPUNS API (PRIMIRE) ---")
            print(f"Status Code: {response.status_code}")
            print(f"Răspunsul Brut (Text): \n{response.text}") # .text este ok pentru debugging
        
        response.raise_for_status()
        
        # Folosim .json() direct
        data = response.json()

        if verbose:
            print(f"DEBUG: Cheile din JSON-ul parsat au fost: {list(data.keys())}")
            print("="*50 + "\n")

        # --- SOLUȚIA CORECTĂ ---
        # Căutăm 'items' ÎN INTERIORUL 'pager'
        items = data.get('pager', {}).get('items', [])
        # --- SFÂRȘIT SOLUȚIE ---


        if not items:
            print(f"Nu s-au găsit date pentru Regiunea {region_id}, Pagina {page_number}.")
            return None
        
        # ACEASTĂ LINIE AR TREBUI SĂ APARĂ ACUM
        print(f"Date găsite: {len(items)} item-uri.")
        
        df = pd.DataFrame(items)
        for col in COLOANE_DORITE:
            if col not in df.columns:
                df[col] = None
        df = df[COLOANE_DORITE]
        
        filename = f'ceccar_data_reg_{region_id}_pag_{page_number}.xlsx'
        filepath = os.path.join(TEMP_DIR, filename)
        
        df.to_excel(filepath, index=False, sheet_name=f'Pagina {page_number}')
        
        print(f"Datele au fost salvate cu succes în: {filepath}")
        return filepath
        
    except Exception as e:
        print(f"A apărut o eroare în `extract_single_page`: {e}")
        return None

# --- Bloc de testare ---
if __name__ == '__main__':
    try:
        import requests
        import pandas
        import openpyxl
    except ImportError:
        print("EROARE: Asigură-te că ai instalat bibliotecile necesare:")
        print("pip install requests pandas openpyxl")
        sys.exit(1)
    
    # --- OPȚIUNEA 1: Extrage TOATE paginile ---
    TEST_REGION_ID_ALL = 2 # Arad
    extract_all_pages_for_region(TEST_REGION_ID_ALL)

    # --- OPȚIUNEA 2: Extrage o SINGURĂ pagină (Activă implicit) ---
    TEST_REGION_ID_SINGLE = 4  # Bacau
    TEST_PAGE_NUMBER_SINGLE = 5 # Pagina 5
    
    #extract_single_page(TEST_REGION_ID_SINGLE, TEST_PAGE_NUMBER_SINGLE, verbose=True)