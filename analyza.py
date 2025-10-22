import pandas as pd
import numpy as np
import os

# --- DEBUGGING ---
# Pokud je True, vytvoří se podrobný log soubor 'debug_vypoctu.txt'
DEBUG = True
DEBUG_FILE = 'debug_vypoctu.txt'

# --- KONFIGURACE ---
# Tento slovník se nyní používá pouze pro získání seznamu názvů krajů.
# Skutečné počty mandátů se vypočítají dynamicky podle zákona.
KRAJE_NAZVY = [
    'Hlavni_mesto_Praha', 'Stredocesky_kraj', 'Jihocesky_kraj',
    'Plzensky_kraj', 'Karlovarsky_kraj', 'Ustecky_kraj',
    'Liberecky_kraj', 'Kralovehradecky_kraj', 'Pardubicky_kraj',
    'Kraj_Vysocina', 'Jihomoravsky_kraj', 'Olomoucky_kraj',
    'Zlinsky_kraj', 'Moravskoslezsky_kraj'
]
VOLEBNI_KLAUZULE = 5.0
NAZEV_VYSTUPNIHO_SOUBORU = 'volebni_analyza_2025_finalni_opravena.xlsx'

MAPOVANI_NAZVU_STRAN = {
    'Svoboda a př. demokracie (SPD)': 'SPD',
    'SPOLU (ODS, KDU-ČSL, TOP 09)': 'SPOLU',
    'Česká pirátská strana': 'Piráti',
    'Motoristé sobě': 'AUTO',
    'ANO 2011': 'ANO',
    'STAROSTOVÉ A NEZÁVISLÍ': 'STAN'
}

# --- FUNKCE PRO DEBUGGOVÁNÍ ---
def initialize_debug_log():
    """Vytvoří nebo vymaže log soubor na začátku běhu."""
    if DEBUG:
        if os.path.exists(DEBUG_FILE):
            os.remove(DEBUG_FILE)
        with open(DEBUG_FILE, 'a', encoding='utf-8') as f:
            f.write("--- NOVÝ BĚH SKRIPTU ---\n\n")

def debug_log(message, level=0):
    """Zapíše zprávu do logovacího souboru."""
    if DEBUG:
        with open(DEBUG_FILE, 'a', encoding='utf-8') as f:
            f.write("    " * level + str(message) + "\n")

# --- VÝPOČETNÍ FUNKCE ---

def spocitej_mandaty_pro_kraje():
    """
    Klíčová funkce, která dynamicky vypočítá, kolik mandátů připadá
    jednotlivým krajům na základě celkového počtu odevzdaných hlasů.
    (Podle § 48 zákona č. 247/1995 Sb.)
    """
    debug_log("--- FÁZE 0: VÝPOČET POČTU MANDÁTŮ PRO KAŽDÝ KRAJ ---", level=0)
    hlasy_v_krajich = {}
    for nazev_kraje in KRAJE_NAZVY:
        try:
            df_kraj = pd.read_csv(f'vysledky_ps2025_{nazev_kraje}.csv', sep=';')
            hlasy_v_krajich[nazev_kraje] = df_kraj['hlasy_celkem'].sum()
        except FileNotFoundError:
            debug_log(f"CHYBA: Soubor pro kraj {nazev_kraje} nebyl nalezen!", level=1)
            return None
    
    df_hlasy = pd.Series(hlasy_v_krajich, name="hlasy_celkem_kraj")
    debug_log("Celkem odevzdaných hlasů v krajích (vstup pro výpočet):", level=1)
    debug_log(df_hlasy.to_string(), level=2)
    
    celkove_hlasy_cr = df_hlasy.sum()
    debug_log(f"\nCelkem hlasů v ČR (bez zahraničí): {celkove_hlasy_cr}", level=1)
    
    republikove_mandatove_cislo_float = celkove_hlasy_cr / 200
    republikove_mandatove_cislo = round(republikove_mandatove_cislo_float)
    debug_log(f"Republikové mandátové číslo (float): {celkove_hlasy_cr} / 200 = {republikove_mandatove_cislo_float}", level=1)
    debug_log(f"Republikové mandátové číslo (zaokrouhleno): {republikove_mandatove_cislo}", level=1)
    
    df_kraje_mandaty = pd.DataFrame(df_hlasy)
    df_kraje_mandaty['mandaty_kolo1'] = df_kraje_mandaty['hlasy_celkem_kraj'] // republikove_mandatove_cislo
    df_kraje_mandaty['zbytek'] = df_kraje_mandaty['hlasy_celkem_kraj'] % republikove_mandatove_cislo
    
    debug_log("\nVýpočet mandátů pro kraje - 1. krok (dělení RČ):", level=1)
    debug_log(df_kraje_mandaty.to_string(), level=2)
    
    prideleno_mandatu_kolo1 = int(df_kraje_mandaty['mandaty_kolo1'].sum())
    zbyva_pridelit = 200 - prideleno_mandatu_kolo1
    debug_log(f"\nPřiděleno mandátů v 1. kroku: {prideleno_mandatu_kolo1}", level=1)
    debug_log(f"Zbývá přidělit podle největších zbytků: {zbyva_pridelit}", level=1)

    df_kraje_mandaty = df_kraje_mandaty.sort_values(by='zbytek', ascending=False)
    debug_log("\nKraje seřazené podle největších zbytků:", level=1)
    debug_log(df_kraje_mandaty.to_string(), level=2)

    df_kraje_mandaty['mandaty_kolo2'] = 0
    for i in range(zbyva_pridelit):
        kraj_k_prideleni = df_kraje_mandaty.index[i]
        df_kraje_mandaty.loc[kraj_k_prideleni, 'mandaty_kolo2'] = 1
    
    df_kraje_mandaty['mandaty_final'] = (df_kraje_mandaty['mandaty_kolo1'] + df_kraje_mandaty['mandaty_kolo2']).astype(int)
    
    debug_log("\nFINÁLNÍ ROZDĚLENÍ MANDÁTŮ KRAJŮM:", level=1)
    debug_log(df_kraje_mandaty[['mandaty_final']].to_string(), level=2)
    
    finalni_mandaty_dict = df_kraje_mandaty['mandaty_final'].to_dict()
    debug_log(f"\nSlovník mandátů pro další výpočet: {finalni_mandaty_dict}", level=1)
    
    return finalni_mandaty_dict


def prvni_skrutinium_imperiali(vysledky_stran_kraj, pocet_mandatu_kraj):
    hlasy_celkem_kraj = sum(vysledky_stran_kraj.values())
    if hlasy_celkem_kraj == 0 or pocet_mandatu_kraj == 0:
        return {s: 0 for s in vysledky_stran_kraj}, {s: h for s, h in vysledky_stran_kraj.items()}, {s: h for s, h in vysledky_stran_kraj.items()}, 0

    volebni_cislo = round(hlasy_celkem_kraj / (pocet_mandatu_kraj + 2))
    mandaty, zbytky_kolo2, zbytky_pridel = {}, {}, {}

    for s, h in vysledky_stran_kraj.items():
        if volebni_cislo > 0:
            mandaty[s] = int(h // volebni_cislo)
            zbytky_pridel[s] = h % volebni_cislo
            zbytky_kolo2[s] = h if mandaty[s] == 0 else h - (mandaty[s] * volebni_cislo)
        else:
            mandaty[s], zbytky_kolo2[s], zbytky_pridel[s] = 0, h, h
            
    return mandaty, zbytky_kolo2, zbytky_pridel, volebni_cislo

def druhe_skrutinium_kompletni(zbytky_hlasu_celkem, nerozdelene_mandaty_celkem):
    if nerozdelene_mandaty_celkem <= 0: return {}, 0
    celkem_zbytkovych_hlasu = sum(zbytky_hlasu_celkem.values())
    if celkem_zbytkovych_hlasu == 0: return {}, 0
        
    volebni_cislo = round(celkem_zbytkovych_hlasu / (nerozdelene_mandaty_celkem + 1))
    mandaty, zbytky_poradi = {}, {}
    
    for s, h in zbytky_hlasu_celkem.items():
        if volebni_cislo > 0:
            mandaty[s] = int(h // volebni_cislo)
            zbytky_poradi[s] = h - (mandaty[s] * volebni_cislo)
        else:
            mandaty[s], zbytky_poradi[s] = 0, h

    prideleno_dosud = sum(mandaty.values())
    zbyva = int(nerozdelene_mandaty_celkem - prideleno_dosud)
    
    serazene_zbytky = sorted(zbytky_poradi.items(), key=lambda i: i[1], reverse=True)
    for i in range(zbyva):
        if i < len(serazene_zbytky):
            mandaty[serazene_zbytky[i][0]] += 1
            
    return mandaty, volebni_cislo

# --- HLAVNÍ FUNKCE ---

def analyzuj_vysledky():
    try:
        # FÁZE 0: Dynamický výpočet mandátů pro kraje
        pocty_mandatu_vypoctene = spocitej_mandaty_pro_kraje()
        if pocty_mandatu_vypoctene is None: return

        df_celkem = pd.read_csv('vysledky_ps2025_Celkem.csv', sep=';')
        strany_nad_5_procent = df_celkem[df_celkem['hlasy_procenta'] >= VOLEBNI_KLAUZULE]
        seznam_uspesnych_stran = list(strany_nad_5_procent['nazev_strany'])
        
        vysledky_1_skrutinia, zbytky_kolo2_dict, zbytky_pridel_dict = {}, {}, {}
        použité_hlasy = {s: 0 for s in seznam_uspesnych_stran}
        
        # FÁZE 1: První skrutinium v krajích
        for nazev, pocet_mandatu in pocty_mandatu_vypoctene.items():
            df_kraj = pd.read_csv(f'vysledky_ps2025_{nazev}.csv', sep=';')
            df_kraj_uspesne = df_kraj[df_kraj['nazev_strany'].isin(seznam_uspesnych_stran)]
            hlasy_dict = dict(zip(df_kraj_uspesne['nazev_strany'], df_kraj_uspesne['hlasy_celkem']))
            
            mandaty, z_k2, z_pr, kvc = prvni_skrutinium_imperiali(hlasy_dict, pocet_mandatu)
            
            vysledky_1_skrutinia[nazev], zbytky_kolo2_dict[nazev], zbytky_pridel_dict[nazev] = mandaty, z_k2, z_pr
            for s, m in mandaty.items():
                použité_hlasy[s] += m * kvc
        
        # FÁZE 2: Druhé skrutinium na celostátní úrovni
        df_mandaty_1 = pd.DataFrame(vysledky_1_skrutinia).fillna(0).astype(int)
        nerozdeleno = 200 - df_mandaty_1.sum().sum()
        df_zbytky_k2 = pd.DataFrame(zbytky_kolo2_dict).fillna(0)
        
        mandaty_z_2, rvc = druhe_skrutinium_kompletni(df_zbytky_k2.sum(axis=1).to_dict(), nerozdeleno)

        # FÁZE 3: Přidělení mandátů z 2. kola zpět do krajů
        fin_mandaty = df_mandaty_1.copy()
        df_zbytky_pridel = pd.DataFrame(zbytky_pridel_dict).fillna(0)
        for strana, pocet in mandaty_z_2.items():
            if pocet > 0:
                for kraj in df_zbytky_pridel.loc[strana].nlargest(pocet).index:
                    fin_mandaty.loc[strana, kraj] += 1

        # ZPRACOVÁNÍ A VÝPIS VÝSLEDKŮ
        df_vysledek = fin_mandaty.transpose()
        df_vysledek = df_vysledek.fillna(0).astype(int).rename(columns=MAPOVANI_NAZVU_STRAN)
        df_vysledek = df_vysledek.reindex(columns=['SPD', 'SPOLU', 'Piráti', 'AUTO', 'ANO', 'STAN'], fill_value=0)
        df_vysledek['Celkem ČR'] = df_vysledek.sum(axis=1)
        df_vysledek.loc['Celkem mandátů'] = df_vysledek.sum()
        df_vysledek = df_vysledek.rename(index={n: n.replace('_', ' ').replace('Hlavni mesto', 'Hlavní město').replace(' kraj', 'ý kraj').replace('Kraj V', 'Kraj V') for n in df_vysledek.index})
        
        print("--- FINÁLNÍ VÝSLEDNÁ TABULKA ---")
        print(df_vysledek.to_string())
        debug_log("\n--- FINÁLNÍ VÝSLEDNÁ TABULKA ---\n" + df_vysledek.to_string(), level=0)
        
        # Analýza hlasů
        analyza = []
        for s in seznam_uspesnych_stran:
            celkem_h = df_celkem[df_celkem['nazev_strany'] == s]['hlasy_celkem'].iloc[0]
            pouzito_2 = mandaty_z_2.get(s, 0) * rvc
            analyza.append({
                'Strana': MAPOVANI_NAZVU_STRAN.get(s, s),
                'Celkem hlasů': celkem_h,
                'Hlasy použité v 1. skrutiniu': int(použité_hlasy[s]),
                'Hlasy použité v 2. skrutiniu': int(pouzito_2),
                'Přeteklé / nevyužité hlasy': int(celkem_h - použité_hlasy[s] - pouzito_2)
            })
        df_analyza = pd.DataFrame(analyza)
        df_analyza.loc['Celkem'] = df_analyza.sum(numeric_only=True)
        propadle = df_celkem[df_celkem['hlasy_procenta'] < VOLEBNI_KLAUZULE]['hlasy_celkem'].sum()
        df_analyza.loc['Propadlé hlasy (strany pod 5 %)'] = [None, propadle, None, None, None]
        df_analyza.loc['Celkem odevzdaných hlasů'] = [None, df_celkem['hlasy_celkem'].sum(), None, None, None]
        
        print("\n--- Analýza využití hlasů ---")
        print(df_analyza.to_string())
        
        with pd.ExcelWriter(NAZEV_VYSTUPNIHO_SOUBORU) as writer:
            df_vysledek.to_excel(writer, sheet_name='Pridelene_mandaty', index_label="Volební kraj")
            df_analyza.to_excel(writer, sheet_name='Analyza_hlasu', index=False)
        print(f"\n✔ Analýza dokončena. Výsledky uloženy do: '{NAZEV_VYSTUPNIHO_SOUBORU}'")

    except Exception as e:
        print(f"\nDošlo k závažné chybě: {e}")
        debug_log(f"\n!!!!!! SKRIPT SELHAL S CHYBOU !!!!!!\n{e}")

if __name__ == '__main__':
    initialize_debug_log()
    analyzuj_vysledky()