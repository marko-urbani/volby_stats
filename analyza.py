import pandas as pd
import numpy as np
import os
from collections import Counter

# --- KONFIGURACE ---
KRAJE_NAZVY = [
    'Hlavni_mesto_Praha', 'Stredocesky_kraj', 'Jihocesky_kraj', 'Plzensky_kraj',
    'Karlovarsky_kraj', 'Ustecky_kraj', 'Liberecky_kraj', 'Kralovehradecky_kraj',
    'Pardubicky_kraj', 'Kraj_Vysocina', 'Jihomoravsky_kraj', 'Olomoucky_kraj',
    'Zlinsky_kraj', 'Moravskoslezsky_kraj'
]
VOLEBNI_KLAUZULE = 5.0
NAZEV_VYSTUPNIHO_SOUBORU = 'volebni_analyza_2025_kompletni.xlsx'

MAPOVANI_NAZVU_STRAN = {
    'Svoboda a př. demokracie (SPD)': 'SPD',
    'SPOLU (ODS, KDU-ČSL, TOP 09)': 'SPOLU',
    'Česká pirátská strana': 'Piráti',
    'Motoristé sobě': 'AUTO',
    'ANO 2011': 'ANO',
    'STAROSTOVÉ A NEZÁVISLÍ': 'STAN'
}

# --- POMOCNÉ FUNKCE ---

def write_section(writer, sheet_name, title, text, dataframe, start_row):
    """Zapíše sekci (nadpis, text a tabulku) na jeden list Excelu s xlsxwriter."""
    dataframe.to_excel(writer, sheet_name=sheet_name, startrow=start_row + 9, header=False)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    title_format = workbook.add_format({'bold': True, 'font_size': 14, 'bottom': 1})
    text_format = workbook.add_format({'text_wrap': True, 'valign': 'top', 'italic': True})
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'border': 1, 'align': 'center', 'bg_color': '#F2F2F2'})

    worksheet.merge_range(start_row, 0, start_row, 15, title, title_format)
    worksheet.merge_range(start_row + 2, 0, start_row + 6, 15, text, text_format)
    
    current_write_row = start_row + 8
    
    for col_num, value in enumerate(dataframe.columns.values):
        worksheet.write(current_write_row, col_num + 1, value, header_format)
    if dataframe.index.name:
        worksheet.write(current_write_row, 0, dataframe.index.name, header_format)
    
    return current_write_row + len(dataframe) + 5

def autofit_columns(worksheet, dataframe, index=False):
    """Přizpůsobí šířku sloupců na daném listu."""
    offset = 1 if index else 0
    for idx, col in enumerate(dataframe.columns):
        series = dataframe[col]
        max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 2
        worksheet.set_column(idx + offset, idx + offset, max_len)
    if index:
         max_len = max((dataframe.index.astype(str).map(len).max(), len(str(dataframe.index.name)))) + 2
         worksheet.set_column(0, 0, max_len)

# --- VÝPOČETNÍ FUNKCE ---
def spocitej_mandaty_pro_kraje():
    hlasy = {n: pd.read_csv(f'vysledky_ps2025_{n}.csv', sep=';')['hlasy_celkem'].sum() for n in KRAJE_NAZVY}
    df = pd.DataFrame.from_dict(hlasy, orient='index', columns=['celkem hlasu'])
    rmc = round(df['celkem hlasu'].sum() / 200)
    df['mandaty krok 1'] = df['celkem hlasu'] // rmc; df['zbytek'] = df['celkem hlasu'] % rmc
    zbyva = 200 - int(df['mandaty krok 1'].sum())
    df = df.sort_values(by='zbytek', ascending=False)
    df['mandaty krok 2'] = [1] * zbyva + [0] * (len(df) - zbyva)
    df['mandaty final'] = (df['mandaty krok 1'] + df['mandaty krok 2']).astype(int)
    df.loc['CELKEM'] = df.sum(numeric_only=True)
    return df, df.loc[df.index != 'CELKEM', 'mandaty final'].to_dict(), rmc

def prvni_skrutinium_imperiali(hlasy, pocet_mandatu):
    celkem = sum(hlasy.values()); kvc = round(celkem / (pocet_mandatu + 2)) if pocet_mandatu > -2 else 0
    m, z2, z3 = {}, {}, {}
    for s, h in hlasy.items():
        if kvc > 0: m[s]=int(h//kvc); z3[s]=h%kvc; z2[s]=h if m[s]==0 else h-(m[s]*kvc)
        else: m[s], z2[s], z3[s] = 0, h, h
    return m, z2, z3, kvc

def druhe_skrutinium_kompletni(zbytky_celkem, nerozdeleno):
    if nerozdeleno <= 0: return {}, 0, pd.DataFrame()
    celkem = sum(zbytky_celkem.values()); rvc = round(celkem / (nerozdeleno + 1)) if nerozdeleno > -1 else 0
    mandaty, zbytky = {}, {}
    for s, h in zbytky_celkem.items():
        if rvc > 0: mandaty[s]=int(h//rvc); zbytky[s]=h-(mandaty[s]*rvc)
        else: mandaty[s], zbytky[s] = 0, h
    zbyva = int(nerozdeleno - sum(mandaty.values()))
    serazene = sorted(zbytky.items(), key=lambda i: i[1], reverse=True)
    for i in range(zbyva):
        if i < len(serazene): mandaty[serazene[i][0]] += 1
    df = pd.DataFrame.from_dict(zbytky_celkem, orient='index', columns=['celkem zbytkovych hlasu'])
    df['mandaty delenim'] = df['celkem zbytkovych hlasu'].apply(lambda x: int(x//rvc) if rvc>0 else 0)
    df['zbytek pro poradi'] = df['celkem zbytkovych hlasu'] - df['mandaty delenim']*rvc
    df = df.sort_values('zbytek pro poradi', ascending=False)
    df['mandaty dle zbytku'] = [1]*zbyva + [0]*(len(df)-zbyva)
    df['mandaty celkem f2'] = (df['mandaty delenim'] + df['mandaty dle zbytku']).astype(int)
    df.columns = [c.replace('_', ' ') for c in df.columns]
    df.loc['CELKEM'] = df.sum(numeric_only=True)
    return mandaty, rvc, df

# --- HLAVNÍ FUNKCE ---
def analyzuj_vysledky():
    try:
        df_mandaty_kraje, pocty_mandatu, rmc_hodnota = spocitej_mandaty_pro_kraje()
        df_celkem = pd.read_csv('vysledky_ps2025_Celkem.csv', sep=';')
        uspesne_strany = list(df_celkem[df_celkem['hlasy_procenta'] >= VOLEBNI_KLAUZULE]['nazev_strany'])
        
        m1_dict, z2_dict, z3_dict, kvc_dict, hlasy_dict_uspesne = {}, {}, {}, {}, {}
        for nazev, pocet in pocty_mandatu.items():
            df_k = pd.read_csv(f'vysledky_ps2025_{nazev}.csv', sep=';'); hlasy = dict(zip(df_k['nazev_strany'], df_k['hlasy_celkem']))
            hlasy_uspesnych = {s: hlasy.get(s, 0) for s in uspesne_strany}; hlasy_dict_uspesne[nazev] = hlasy_uspesnych
            m, z2, z3, kvc = prvni_skrutinium_imperiali(hlasy_uspesnych, pocet)
            m1_dict[nazev], z2_dict[nazev], z3_dict[nazev], kvc_dict[nazev] = m, z2, z3, kvc
        df_hlasy_uspesne = pd.DataFrame(hlasy_dict_uspesne).fillna(0).astype(int)
        df_m1 = pd.DataFrame(m1_dict).fillna(0).astype(int); df_z2 = pd.DataFrame(z2_dict).fillna(0).astype(int); df_z3 = pd.DataFrame(z3_dict).fillna(0).astype(int)
        
        mandaty_nerozdeleno_f1 = 200 - df_m1.sum().sum()
        mandaty_f2, rvc_hodnota, df_f2_vypocet = druhe_skrutinium_kompletni(df_z2.sum(axis=1).to_dict(), mandaty_nerozdeleno_f1)
        
        fin_mandaty = df_m1.copy()
        for s, p in mandaty_f2.items():
            if p > 0:
                for kraj in df_z3.loc[s].nlargest(p).index: fin_mandaty.loc[s, kraj] += 1
        df_vysledek = fin_mandaty.transpose().rename(columns=MAPOVANI_NAZVU_STRAN).reindex(columns=list(MAPOVANI_NAZVU_STRAN.values()), fill_value=0)
        df_vysledek['Celkem mandátů v kraji'] = df_vysledek.sum(axis=1); df_vysledek.loc['Celkem mandátů strany'] = df_vysledek.sum()

        pouzite_f1_total = sum(df_m1.loc[s, k] * kvc_dict[k] for s in df_m1.index for k in df_m1.columns)
        df_f2_cisty = df_f2_vypocet.drop('CELKEM', errors='ignore')
        pouzite_f2_total = (
            (df_f2_cisty['mandaty delenim'] * rvc_hodnota).sum()
            + df_f2_cisty.loc[df_f2_cisty['mandaty dle zbytku'] > 0, 'zbytek pro poradi'].sum()
        )

        df_nevyuzite_final = df_f2_vypocet.loc[(df_f2_vypocet.index != 'CELKEM') & (df_f2_vypocet['mandaty dle zbytku'] == 0), ['zbytek pro poradi']].rename(columns={'zbytek pro poradi': 'Nevyužité hlasy (zbytky po 2. skrutiniu)'})
        if not df_nevyuzite_final.empty:
            df_nevyuzite_final.loc['CELKEM'] = df_nevyuzite_final.sum()
            nevyuzite_total = df_nevyuzite_final.loc['CELKEM'].iloc[0]
        else: nevyuzite_total = 0
            
        propadle_total = df_celkem[df_celkem['hlasy_procenta'] < VOLEBNI_KLAUZULE]['hlasy_celkem'].sum()
        celkem_propadlo_nevyuzito = propadle_total + nevyuzite_total; celkem_vsech_hlasu = df_celkem['hlasy_celkem'].sum()
        
        df_rozpis_data = [{'Kategorie': 'Hlasy využité v 1. skrutiniu', 'Počet hlasů': int(pouzite_f1_total), 'Počet mandátů': int(df_m1.sum().sum())}, {'Kategorie': 'Hlasy využité v 2. skrutiniu', 'Počet hlasů': int(pouzite_f2_total), 'Počet mandátů': int(mandaty_nerozdeleno_f1)}, {'Kategorie': 'CELKEM VYUŽITO', 'Počet hlasů': int(pouzite_f1_total + pouzite_f2_total), 'Počet mandátů': 200}, {'Kategorie': ''}, {'Kategorie': 'Hlasy propadlé (strany < 5 %)', 'Počet hlasů': int(propadle_total)}, {'Kategorie': 'Hlasy nevyužité (zbytky úspěšných stran)', 'Počet hlasů': int(nevyuzite_total)}, {'Kategorie': 'CELKEM PROPADLO A NEVYUŽITO', 'Počet hlasů': int(celkem_propadlo_nevyuzito)}, {'Kategorie': 'CELKEM ODEVZDANÝCH HLASŮ', 'Počet hlasů': int(celkem_vsech_hlasu)}, {'Kategorie': 'Podíl propadlých a nevyužitých hlasů', 'Počet hlasů': f"{round((celkem_propadlo_nevyuzito / celkem_vsech_hlasu) * 100, 2)} %"}]
        df_rozpis = pd.DataFrame(df_rozpis_data)
        hlasy_numeric = pd.to_numeric(df_rozpis['Počet hlasů'], errors='coerce'); mandaty_numeric = pd.to_numeric(df_rozpis['Počet mandátů'], errors='coerce')
        df_rozpis['Průměr na 1 mandát'] = np.divide(hlasy_numeric, mandaty_numeric).fillna(0).apply(lambda x: f"{int(x)}" if x != 0 else '')

        ceny_mandatu = []
        for s in df_m1.index:
            for k in df_m1.columns:
                if df_m1.loc[s, k] > 0: ceny_mandatu.extend([kvc_dict[k]] * df_m1.loc[s, k])
        
        df_f2_cisty = df_f2_vypocet.drop('CELKEM', errors='ignore')
        for s, row in df_f2_cisty.iterrows():
            if row['mandaty delenim'] > 0: ceny_mandatu.extend([rvc_hodnota] * row['mandaty delenim'])
            if row['mandaty dle zbytku'] > 0: ceny_mandatu.append(row['zbytek pro poradi'])
            
        pocetnost_cen = Counter(ceny_mandatu)
        df_ceny = pd.DataFrame(pocetnost_cen.items(), columns=['Cena mandátu (počet hlasů)', 'Počet mandátů za tuto cenu'])
        df_ceny = df_ceny.sort_values(by='Cena mandátu (počet hlasů)', ascending=False).reset_index(drop=True)
        df_ceny['Celkem hlasů'] = df_ceny['Cena mandátu (počet hlasů)'] * df_ceny['Počet mandátů za tuto cenu']
        df_ceny = df_ceny[['Cena mandátu (počet hlasů)', 'Počet mandátů za tuto cenu', 'Celkem hlasů']].astype(int)
        df_ceny.index.name = 'Pořadí'; df_ceny.index += 1

        try:
            writer = pd.ExcelWriter(NAZEV_VYSTUPNIHO_SOUBORU, engine='xlsxwriter')
            df_m_k_vystup = df_mandaty_kraje.rename(index={n:n.replace('_',' ') for n in df_mandaty_kraje.index})
            df_hlasy_vystup = df_hlasy_uspesne.rename(index=MAPOVANI_NAZVU_STRAN, columns={n:n.replace('_',' ') for n in df_hlasy_uspesne.columns}); df_hlasy_vystup['CELKEM STRANY']=df_hlasy_vystup.sum(axis=1); df_hlasy_vystup.loc['CELKEM V KRAJI']=df_hlasy_vystup.sum()
            df_m1_vystup = df_m1.rename(index=MAPOVANI_NAZVU_STRAN, columns={n:n.replace('_',' ') for n in df_m1.columns}); df_m1_vystup['CELKEM STRANY']=df_m1_vystup.sum(axis=1); df_m1_vystup.loc['CELKEM V KRAJI']=df_m1_vystup.sum()
            df_z2_vystup = df_z2.rename(index=MAPOVANI_NAZVU_STRAN, columns={n:n.replace('_',' ') for n in df_z2.columns}); df_z2_vystup['CELKEM STRANY']=df_z2_vystup.sum(axis=1); df_z2_vystup.loc['CELKEM V KRAJI']=df_z2_vystup.sum()
            df_f2_vystup = df_f2_vypocet.rename(index=MAPOVANI_NAZVU_STRAN)
            uspesne_mapovane = {k: v for k, v in MAPOVANI_NAZVU_STRAN.items() if k in uspesne_strany}
            df_hlasy_celkem_uspesne = df_celkem[df_celkem['nazev_strany'].isin(uspesne_strany)].set_index('nazev_strany')
            df_mandaty_celkem_uspesne = df_vysledek.loc['Celkem mandátů strany']
            efektivita_data = [{'Strana': map_nazev, 'Celkem hlasů': df_hlasy_celkem_uspesne.loc[orig_nazev, 'hlasy_celkem'], 'Celkem mandátů': int(df_mandaty_celkem_uspesne.get(map_nazev, 0)), 'Průměr hlasů na mandát': int(df_hlasy_celkem_uspesne.loc[orig_nazev, 'hlasy_celkem'] / df_mandaty_celkem_uspesne.get(map_nazev, 1)) if df_mandaty_celkem_uspesne.get(map_nazev, 0) > 0 else 0} for orig_nazev, map_nazev in uspesne_mapovane.items()]
            df_efektivita = pd.DataFrame(efektivita_data).set_index('Strana')
            df_kvc_vystup = pd.DataFrame.from_dict(kvc_dict, orient='index', columns=['Krajské volební číslo (KVČ)']).rename_axis('Kraj')
            df_kvc_vystup.index = df_kvc_vystup.index.str.replace('_', ' ')
            df_vysledek_vystup = df_vysledek.rename(index={n: n.replace('_', ' ') for n in df_vysledek.index})


            with writer:
                sheet_calc = '00_Postup_vypoctu'
                popisky = {
                    "KROK1": f"Nejprve se rozdělí 200 mandátů mezi kraje dle odevzdaných hlasů. Použije se republikové mandátové číslo (RMČ), které pro tyto volby činí {rmc_hodnota:,} hlasů, a metoda největších zbytků.\nPočet hlasů v kraji se vydělí RMČ. Tím kraj získá první várku mandátů. Zbylé mandáty se pak přidělí krajům s největším zbytkem po tomto dělení.",
                    "KROK2": "Do výpočtu vstupují pouze hlasy pro strany, které na celostátní úrovni překročily 5% hranici. Hlasy pro ostatní strany propadají.",
                    "KROK3": "Pro každý kraj se spočítá unikátní krajské volební číslo (KVČ). Vypočítá se tak, že se součet hlasů úspěšných stran v daném kraji vydělí počtem mandátů pro tento kraj, zvětšeným o 2.",
                    "KROK4": "Počet mandátů pro stranu v kraji se určí tak, že se vezme počet hlasů dané strany v kraji, vydělí se krajským volebním číslem (KVČ) a výsledek se zaokrouhlí dolů na celé číslo.",
                    "KROK5": "Tato tabulka ukazuje hlasy, které nebyly využity v 1. kole a postupují do 2. kola. Podle zákona mají dvojí roli: slouží jak pro výpočet mandátů ve 2. kole, tak pro jejich následné umístění do krajů. V těchto konkrétních výsledcích jsou hodnoty pro oba účely totožné, protože žádná strana v žádném kraji nezůstala v 1. kole bez mandátu.",
                    "KROK6": f"Zde se sečtou všechny hlasy z předchozího kroku a rozdělí se jimi zbývající mandáty. Používá se republikové volební číslo (RVČ), které pro tyto volby činí {rvc_hodnota:,} hlasů, a metoda největších zbytků.",
                    "KROK7": "Toto je finální přehled rozdělení všech 200 mandátů. Tabulka sčítá mandáty získané v 1. skrutiniu s mandáty z 2. skrutinia, které byly na základě největších zbytků hlasů přiděleny do konkrétních krajů.",
                    "KROK8": "Hlasy úspěšných stran, které postoupily do 2. kola, ale ani zde nestačily na zisk mandátu. Jedná se o zbytky po dělení RVČ u stran, kterým nebyl přidělen mandát na základě pořadí zbytků.",
                    "KROK9": "Tato tabulka shrnuje, jak byly všechny odevzdané hlasy využity, kolik jich bylo potřeba na mandáty v jednotlivých fázích a kolik jich celkem propadlo nebo zůstalo nevyužito.",
                    "KROK10": "Tato tabulka ukazuje průměrný počet hlasů, který byl pro každou úspěšnou stranu potřeba k zisku jednoho mandátu. Nižší číslo znamená vyšší efektivitu přeměny hlasů na mandáty.",
                    "KROK11": "Tato tabulka ukazuje přesné 'ceny' všech 200 rozdělených mandátů. Cena mandátu z 1. kola je krajské volební číslo (KVČ). Cena mandátu z 2. kola je buď republikové volební číslo (RVČ), nebo výše zbytkového balíku hlasů, který mandát zajistil. Mandáty se stejnou cenou jsou sečteny."
                }
                all_sections = [
                    (f"Rozdělení mandátů mezi kraje (RMČ = {rmc_hodnota:,})", popisky["KROK1"], df_m_k_vystup),
                    ("Vstupní hlasy pro 1. skrutinium", popisky["KROK2"], df_hlasy_vystup),
                    ("Krajská volební čísla (KVČ)", popisky["KROK3"], df_kvc_vystup),
                    ("Mandáty přidělené v 1. skrutiniu", popisky["KROK4"], df_m1_vystup),
                    ("Hlasy postupující do 2. skrutinia", popisky["KROK5"], df_z2_vystup),
                    (f"Výpočet 2. skrutinia (RVČ = {rvc_hodnota:,})", popisky["KROK6"], df_f2_vystup),
                    ("Finální rozdělení mandátů po obou skrutiniích", popisky["KROK7"], df_vysledek_vystup),
                    ("Finálně nevyužité hlasy úspěšných stran", popisky["KROK8"], df_nevyuzite_final),
                    ("Celková bilance hlasů", popisky["KROK9"], df_rozpis.set_index('Kategorie')),
                    ("Efektivita hlasů (průměrná cena mandátu)", popisky["KROK10"], df_efektivita),
                    ("Detailní přehled cen všech mandátů", popisky["KROK11"], df_ceny)
                ]
                current_row = 1
                for title, text, df in all_sections:
                    current_row = write_section(writer, sheet_calc, title, text, df, current_row)

                worksheet = writer.sheets[sheet_calc]
                worksheet.set_column('A:A', 35)
                worksheet.set_column('B:R', 12)

            print(f"\n✔ Analýza dokončena. Výsledky uloženy do: '{NAZEV_VYSTUPNIHO_SOUBORU}'")
        except PermissionError:
            print(f"\n❌ CHYBA: Nelze uložit soubor '{NAZEV_VYSTUPNIHO_SOUBORU}'. Je pravděpodobně otevřen. Prosím, zavřete jej a spusťte skript znovu.")
        except Exception as e:
            print(f"\n❌ Chyba při ukládání do Excelu: {e}")

    except FileNotFoundError as e:
        print(f"\n❌ CHYBA: Soubor s volebními výsledky nebyl nalezen: {e.filename}. Ujistěte se, že všechny potřebné .csv soubory jsou ve stejné složce jako skript.")
    except Exception as e:
        print(f"\nDošlo k závažné chybě: {e}")

if __name__ == '__main__':
    analyzuj_vysledky()