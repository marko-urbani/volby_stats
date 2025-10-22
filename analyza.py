import pandas as pd
import numpy as np
import os

# --- DEBUGGING ---
DEBUG = True
DEBUG_FILE = 'debug_vypoctu.txt'

# --- KONFIGURACE ---
# Tento slovník se používá pouze pro získání seznamu názvů krajů.
# Skutečné počty mandátů se vypočítají dynamicky.
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
    if DEBUG:
        try:
            if os.path.exists(DEBUG_FILE):
                os.remove(DEBUG_FILE)
            with open(DEBUG_FILE, 'w', encoding='utf-8') as f:
                f.write("--- EXTRÉMNÍ DEBUGGOVACÍ REŽIM ZAHÁJEN ---\n\n")
        except Exception as e:
            print(f"Nelze inicializovat debug soubor: {e}")

def debug_log(message, level=0):
    if DEBUG:
        try:
            with open(DEBUG_FILE, 'a', encoding='utf-8') as f:
                f.write("    " * level + str(message) + "\n")
        except Exception as e:
            print(f"Nelze zapsat do debug souboru: {e}")

# --- VÝPOČETNÍ FUNKCE ---

def spocitej_mandaty_pro_kraje():
    debug_log("--- FÁZE 0: DYNAMICKÝ VÝPOČET POČTU MANDÁTŮ PRO KAŽDÝ KRAJ ---", level=0)
    hlasy_v_krajich = {}
    for nazev_kraje in KRAJE_NAZVY:
        try:
            df_kraj = pd.read_csv(f'vysledky_ps2025_{nazev_kraje}.csv', sep=';')
            hlasy_v_krajich[nazev_kraje] = df_kraj['hlasy_celkem'].sum()
        except FileNotFoundError:
            debug_log(f"CHYBA: Soubor pro kraj {nazev_kraje} nebyl nalezen!", level=1)
            return None
    
    df_hlasy = pd.Series(hlasy_v_krajich, name="hlasy_celkem_kraj")
    celkove_hlasy_cr = df_hlasy.sum()
    republikove_mandatove_cislo = round(celkove_hlasy_cr / 200)
    
    df_kraje = pd.DataFrame(df_hlasy)
    df_kraje['mandaty_krok1'] = df_kraje['hlasy_celkem_kraj'] // republikove_mandatove_cislo
    df_kraje['zbytek'] = df_kraje['hlasy_celkem_kraj'] % republikove_mandatove_cislo
    
    prideleno_krok1 = int(df_kraje['mandaty_krok1'].sum())
    zbyva_pridelit = 200 - prideleno_krok1
    
    df_kraje = df_kraje.sort_values(by='zbytek', ascending=False)
    df_kraje['mandaty_krok2'] = 0
    for i in range(zbyva_pridelit):
        df_kraje.iloc[i, df_kraje.columns.get_loc('mandaty_krok2')] = 1
        
    df_kraje['mandaty_final'] = (df_kraje['mandaty_krok1'] + df_kraje['mandaty_krok2']).astype(int)
    
    debug_log(f"Celkem hlasů v ČR (bez zahraničí): {celkove_hlasy_cr}", level=1)
    debug_log(f"Republikové mandátové číslo: round({celkove_hlasy_cr} / 200) = {republikove_mandatove_cislo}", level=1)
    debug_log("\nDetailní výpočet mandátů krajů:", level=1)
    debug_log(df_kraje.to_string(), level=2)
    debug_log(f"\nFINÁLNÍ SLOVNÍK MANDÁTŮ KRAJŮ PRO DALŠÍ VÝPOČET: {df_kraje['mandaty_final'].to_dict()}", level=1)
    
    return df_kraje['mandaty_final'].to_dict()

def prvni_skrutinium_imperiali(vysledky_stran_kraj, pocet_mandatu_kraj, nazev_kraje):
    debug_log(f"--- F1: Kraj {nazev_kraje} (přiděluje se {pocet_mandatu_kraj} mandátů) ---", level=1)
    hlasy_celkem_kraj = sum(vysledky_stran_kraj.values())
    if hlasy_celkem_kraj == 0 or pocet_mandatu_kraj == 0:
        return {s: 0 for s in vysledky_stran_kraj}, {s: h for s, h in vysledky_stran_kraj.items()}, {s: h for s, h in vysledky_stran_kraj.items()}, 0

    volebni_cislo = round(hlasy_celkem_kraj / (pocet_mandatu_kraj + 2))
    debug_log(f"Krajské volební číslo: round({hlasy_celkem_kraj} / {pocet_mandatu_kraj + 2}) = {volebni_cislo}", level=2)
    mandaty, zbytky_kolo2, zbytky_pridel = {}, {}, {}

    for s, h in vysledky_stran_kraj.items():
        if volebni_cislo > 0:
            mandaty[s] = int(h // volebni_cislo)
            zbytky_pridel[s] = h % volebni_cislo
            zbytky_kolo2[s] = h if mandaty[s] == 0 else h - (mandaty[s] * volebni_cislo)
        else:
            mandaty[s], zbytky_kolo2[s], zbytky_pridel[s] = 0, h, h
        debug_log(f"-> {s[:25]:<25}: {h:7} hlasů -> {mandaty[s]} mand. | Zbytek pro F2: {zbytky_kolo2[s]:6} | Zbytek pro F3: {zbytky_pridel[s]:6}", level=3)
            
    return mandaty, zbytky_kolo2, zbytky_pridel, volebni_cislo

def druhe_skrutinium_kompletni(zbytky_hlasu_celkem, nerozdelene_mandaty_celkem):
    debug_log("--- FÁZE 2: DRUHÉ SKRUTINIUM ---", level=0)
    if nerozdelene_mandaty_celkem <= 0: return {}, 0
    celkem_zbytkovych_hlasu = sum(zbytky_hlasu_celkem.values())
    if celkem_zbytkovych_hlasu == 0: return {}, 0
        
    volebni_cislo = round(celkem_zbytkovych_hlasu / (nerozdelene_mandaty_celkem + 1))
    debug_log(f"Vstup: {nerozdelene_mandaty_celkem} mandátů, {celkem_zbytkovych_hlasu} hlasů", level=1)
    debug_log(f"Republikové volební číslo: round({celkem_zbytkovych_hlasu} / {nerozdelene_mandaty_celkem + 1}) = {volebni_cislo}", level=1)
    
    mandaty, zbytky_poradi = {}, {}
    
    for s, h in zbytky_hlasu_celkem.items():
        if volebni_cislo > 0:
            mandaty[s] = int(h // volebni_cislo)
            zbytky_poradi[s] = h - (mandaty[s] * volebni_cislo)
        else:
            mandaty[s], zbytky_poradi[s] = 0, h
        debug_log(f"-> {s[:25]:<25}: {h:7} hlasů -> {mandaty[s]} mand. | Zbytek: {zbytky_poradi[s]}", level=2)

    prideleno_dosud = sum(mandaty.values())
    zbyva = int(nerozdelene_mandaty_celkem - prideleno_dosud)
    debug_log(f"Přiděleno dělením: {prideleno_dosud}, zbývá podle zbytků: {zbyva}", level=1)
    
    serazene_zbytky = sorted(zbytky_poradi.items(), key=lambda i: i[1], reverse=True)
    for i in range(zbyva):
        if i < len(serazene_zbytky):
            strana = serazene_zbytky[i][0]
            mandaty[strana] += 1
            debug_log(f"Mandát navíc pro {strana} (zbytek: {serazene_zbytky[i][1]})", level=2)
            
    return mandaty, volebni_cislo

def analyzuj_vysledky():
    try:
        pocty_mandatu_vypoctene = spocitej_mandaty_pro_kraje()
        if pocty_mandatu_vypoctene is None: return

        df_celkem = pd.read_csv('vysledky_ps2025_Celkem.csv', sep=';')
        strany_nad_5 = df_celkem[df_celkem['hlasy_procenta'] >= VOLEBNI_KLAUZULE]
        seznam_uspesnych = list(strany_nad_5['nazev_strany'])
        
        mandaty_f1, zbytky_f2, zbytky_f3 = {}, {}, {}
        
        for nazev, pocet in pocty_mandatu_vypoctene.items():
            df_k = pd.read_csv(f'vysledky_ps2025_{nazev}.csv', sep=';')
            hlasy = dict(zip(df_k['nazev_strany'], df_k['hlasy_celkem']))
            hlasy_uspesnych = {s: hlasy.get(s, 0) for s in seznam_uspesnych}
            m, z2, z3, _ = prvni_skrutinium_imperiali(hlasy_uspesnych, pocet, nazev)
            mandaty_f1[nazev], zbytky_f2[nazev], zbytky_f3[nazev] = m, z2, z3
        
        df_m1 = pd.DataFrame(mandaty_f1).fillna(0).astype(int)
        nerozdeleno = 200 - df_m1.sum().sum()
        df_z2 = pd.DataFrame(zbytky_f2).fillna(0)
        
        mandaty_f2, _ = druhe_skrutinium_kompletni(df_z2.sum(axis=1).to_dict(), nerozdeleno)

        fin_mandaty = df_m1.copy()
        df_z3 = pd.DataFrame(zbytky_f3).fillna(0)
        debug_log("\n--- FÁZE 3: PŘIDĚLENÍ MANDÁTŮ ZPĚT DO KRAJŮ ---", level=0)
        for strana, pocet in mandaty_f2.items():
            if pocet > 0:
                debug_log(f"Přiděluji {pocet} mandátů pro {strana}:", level=1)
                nejvetsi_zbytky = df_z3.loc[strana].nlargest(pocet)
                for kraj, zbytek in nejvetsi_zbytky.items():
                    fin_mandaty.loc[strana, kraj] += 1
                    debug_log(f"-> Mandát přidělen do {kraj} (zbytek: {zbytek})", level=2)

        df_vysledek = fin_mandaty.transpose()
        df_vysledek = df_vysledek.fillna(0).astype(int).rename(columns=MAPOVANI_NAZVU_STRAN)
        df_vysledek = df_vysledek.reindex(columns=['SPD', 'SPOLU', 'Piráti', 'AUTO', 'ANO', 'STAN'], fill_value=0)
        df_vysledek['Celkem ČR'] = df_vysledek.sum(axis=1)
        df_vysledek.loc['Celkem mandátů'] = df_vysledek.sum()
        df_vysledek = df_vysledek.rename(index={n: n.replace('_', ' ').replace('Hlavni mesto', 'Hlavní město').replace(' kraj', 'ý kraj').replace('Kraj V', 'Kraj V') for n in df_vysledek.index})
        
        print("--- FINÁLNÍ VÝSLEDNÁ TABULKA ---")
        print(df_vysledek.to_string())
        debug_log("\n" + df_vysledek.to_string(), level=0)
        
        try:
            with pd.ExcelWriter(NAZEV_VYSTUPNIHO_SOUBORU) as writer:
                df_vysledek.to_excel(writer, "Pridelene_mandaty", index_label="Volební kraj")
            print(f"\n✔ Analýza dokončena. Výsledky uloženy do: '{NAZEV_VYSTUPNIHO_SOUBORU}'")
        except PermissionError:
            print(f"\n❌ CHYBA: Soubor '{NAZEV_VYSTUPNIHO_SOUBORU}' je pravděpodobně otevřen.")
            print("Prosím, zavřete jej a spusťte skript znovu.")
        except Exception as e:
            print(f"\n❌ Chyba při ukládání do Excelu: {e}")

    except Exception as e:
        print(f"\nDošlo k závažné chybě: {e}")
        debug_log(f"\n!!!!!! SKRIPT SELHAL S CHYBOU !!!!!!\n{type(e).__name__}: {e}")

if __name__ == '__main__':
    initialize_debug_log()
    analyzuj_vysledky()