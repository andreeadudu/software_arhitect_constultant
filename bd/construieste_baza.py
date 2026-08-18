"""
================================================================================
 CONSTRUIRE BAZA DE DATE EUROSTAT — 2014-2024, nivel de tara (EU27)
 Lucrare de disertatie: nuptialitate EU27, model spatial (SARAR)
 PAS 2 (partea de extragere). FARA log / standardizare / imputare — date brute.
================================================================================

Rezultat: un fisier Excel cu CATE O FOAIE PE VARIABILA. Fiecare foaie are:
   cod_tara | tara | 2014 | 2015 | ... | 2024
si contine toate cele 27 de tari EU27 (chiar daca unele au valori lipsa pe
ultimii ani — le lasam goale intentionat, ca sa se vada exact unde-s gaurile).

Prima foaie ("Info") listeaza fiecare variabila cu codul, filtrele si notele.

Rulare:
    pip install eurostat pandas openpyxl
    python construieste_baza.py
================================================================================
"""

import re
import sys
import pandas as pd
import numpy as np

try:
    import eurostat
except ImportError:
    sys.exit("Lipseste pachetul 'eurostat'. Ruleaza: pip install eurostat")


# ------------------------------------------------------------------------------
# CONFIGURARE
# ------------------------------------------------------------------------------
ANI = list(range(2014, 2025))          # 2014 ... 2024 inclusiv
OUTPUT_XLSX = "baza_date_2014_2024.xlsx"

# EU27 + nume in romana (Eurostat foloseste 'EL' pt Grecia, nu 'GR').
NUME_TARI = {
    "AT": "Austria", "BE": "Belgia", "BG": "Bulgaria", "HR": "Croatia",
    "CY": "Cipru", "CZ": "Cehia", "DE": "Germania", "DK": "Danemarca",
    "EE": "Estonia", "ES": "Spania", "FI": "Finlanda", "FR": "Franta",
    "EL": "Grecia", "HU": "Ungaria", "IE": "Irlanda", "IT": "Italia",
    "LT": "Lituania", "LU": "Luxemburg", "LV": "Letonia", "MT": "Malta",
    "NL": "Tarile de Jos", "PL": "Polonia", "PT": "Portugalia", "RO": "Romania",
    "SE": "Suedia", "SI": "Slovenia", "SK": "Slovacia",
}
EU27 = list(NUME_TARI.keys())

# --- Variabile care se descarca DIRECT ca o singura serie per tara ---
# (acronim -> cod, filtre, descriere)
DIRECTE = {
    "Mar_rate":        ("demo_nind",   {"indic_de": "GNUPRT"},   "Rata bruta de nuptialitate (la 1000 loc.)"),
    "Mar_rate_prim_F": ("demo_nind",   {"indic_de": "FMAR1CUM"}, "Rata totala a primelor casatorii - femei"),
    "Mar_rate_prim_M": ("demo_nind",   {"indic_de": "MMAR1CUM"}, "Rata totala a primelor casatorii - barbati"),
    "Age_first_mar_F": ("demo_nind",   {"indic_de": "FAGEMAR1"}, "Varsta medie la prima casatorie - femei"),
    "Age_first_mar_M": ("demo_nind",   {"indic_de": "MAGEMAR1"}, "Varsta medie la prima casatorie - barbati"),
    "Marriages_raw":   ("demo_nind",   {"indic_de": "MARRIAGE"}, "Numar absolut de casatorii (ingredient)"),

    "GDP_cap":   ("tec00114",     {},                                                 "PIB/cap, PPS, index EU27=100"),
    "Une_rate":  ("une_rt_a",     {"sex": "T", "age": "Y15-74", "unit": "PC_ACT"},    "Rata somajului (% pop. activa)"),
    "Emp_rate":  ("lfsi_emp_a",   {"sex": "T", "age": "Y20-64", "unit": "PC_POP", "indic_em": "EMP_LFS"}, "Rata de ocupare totala 20-64"),
    "Emp_rate_F":("lfsi_emp_a",   {"sex": "F", "age": "Y20-64", "unit": "PC_POP", "indic_em": "EMP_LFS"}, "Rata de ocupare femei 20-64"),
    "Emp_rate_M":("lfsi_emp_a",   {"sex": "M", "age": "Y20-64", "unit": "PC_POP", "indic_em": "EMP_LFS"}, "Rata de ocupare barbati 20-64"),

    "Ter_educ":  ("edat_lfse_03", {"sex": "T", "age": "Y25-64", "isced11": "ED5-8", "unit": "PC"}, "Pondere 25-64 cu educatie tertiara"),
    "NEET":      ("tesem150",     {"sex": "T", "age": "Y15-29"},                      "Rata NEET 15-29"),
    "Tot_fer_r": ("tps00199",     {},                                                 "Rata totala de fertilitate"),
    "Age_leave": ("yth_demo_030", {"sex": "T"},                                       "Varsta medie la parasirea casei parintesti"),

    "Pov_risk":  ("tps00184",     {"sex": "T", "age": "TOTAL", "unit": "PC"},        "Rata riscului de saracie (prag 60% mediana - tabel titlu AROP)"),
    "Mat_depriv":("ilc_mdsd07",   {"age": "TOTAL", "sex": "T", "unit": "PC"},         "Deprivare materiala si sociala severa"),
    "Subj_pov":  ("ilc_sbjp05",   {"deg_urb": "DEG1"},                                "Saracie subiectiva (orase - DEG1)"),

    "Net_migr_rate": ("demo_gind", {"indic_de": "CNMIGRATRT"},                        "Rata bruta a migratiei nete (la 1000 loc.)"),
    "Liv_parent":    ("ilc_lvps08", {"age": "Y18-34", "sex": "T", "unit": "PC"},      "Pondere tineri 18-34 care locuiesc cu parintii"),
}

# Note per variabila (pt foaia Info); cheile lipsa raman fara nota.
NOTE = {
    "Mar_rate_prim_F": "Acoperire slaba pe ultimii ani (multe tari nu raporteaza). Candidat dependenta, dar gappy.",
    "Mar_rate_prim_M": "Idem - acoperire slaba recenta.",
    "Age_first_mar_F": "Idem - acoperire slaba recenta.",
    "Age_first_mar_M": "Idem - acoperire slaba recenta.",
    "Mar_rate": "Candidatul dependenta recomandat (aproape complet). Lipsa 2024: BE,CY,DK,FR,IE; CY se opreste in 2019.",
    "Mat_depriv": "Serie porneste in 2014, acopera exact fereastra.",
    "Subj_pov": "[confirmat] DEG1='Cities' in clasificarea Eurostat DEGURBA — corespunde definitiei 'saracie subiectiva in orase' din tabelul profesoarei. Nu exista categorie TOTAL in acest dataset.",
    "Emp_rate": "[confirmat] indic_em='EMP_LFS' (rata de ocupare) — nu 'ACT' (rata de activitate, alt concept). unit='PC_POP', nu 'PC'.",
    "Emp_rate_F": "Idem Emp_rate.",
    "Emp_rate_M": "Idem Emp_rate.",
    "Pov_risk": "[CORECTAT de 2 ori] ilc_li02 avea codificare de praguri ambigua (rskpovth A/B x statinfo MED/MEAN) care dadea coloana goala. Am trecut pe tps00184 = tabelul-titlu AROP, prag fix 60% mediana, o singura serie. Daca nici asta nu merge, scriptul opreste cu EROARE si listeaza dimensiunile reale.",
    "Liv_parent": "[corectat] Am schimbat dataset-ul din ilc_lvps09 (defalcat DOAR pe status ocupational, fara total) in ilc_lvps08 (are categoria total sex=T). age='Y18-34' conform cerintei.",
}


# ------------------------------------------------------------------------------
# UTILITARE (format wide al pachetului 'eurostat' din Python)
# ------------------------------------------------------------------------------
def find_geo_col(df):
    for c in df.columns:
        if "geo" in str(c).lower():
            return c
    raise KeyError("Nu am gasit coloana 'geo'.")

def find_year_cols(df):
    return [c for c in df.columns if re.fullmatch(r"\d{4}", str(c).strip())]

def dimension_cols(df, geo_col, year_cols):
    return [c for c in df.columns if c != geo_col and c not in year_cols]

def apply_filters(df, filters, dim_cols, acronim):
    colmap = {c: str(c).split("\\")[0].lower() for c in dim_cols}
    inv = {}
    for real, short in colmap.items():
        inv.setdefault(short, real)
    out = df.copy()
    for key, val in filters.items():
        real_col = inv.get(key.lower())
        if real_col is None:
            print(f"   [!] {acronim}: dimensiunea '{key}' nu exista -> filtru ignorat.")
            continue
        available = set(out[real_col].dropna().unique())
        if val not in available:
            print(f"   [!] {acronim}: valoarea '{val}' lipseste pe '{key}'. Disponibile: {sorted(available)}")
            continue
        out = out[out[real_col] == val]
    return out


def fetch_series(code, filters, acronim, prefer=None):
    """Descarca un dataset si il reduce la o tabela geo(EU27) x ani(2014-2024).
    prefer: daca dupa filtrare raman mai multe randuri per tara, pastreaza randul
    care are aceasta valoare exacta pe orice dimensiune (ex. 'EMP')."""
    df = eurostat.get_data_df(code)
    if df is None or df.empty:
        raise ValueError(f"dataset gol pentru {code}")

    geo_col = find_geo_col(df)
    year_cols = find_year_cols(df)
    dim_cols = dimension_cols(df, geo_col, year_cols)

    # inventar dimensiuni (pt diagnostic daca ceva iese gol)
    inventar = "; ".join(
        f"{str(c).split(chr(92))[0]}={sorted(map(str, df[c].dropna().unique()))[:15]}"
        for c in dim_cols
    )

    df = apply_filters(df, filters, dim_cols, acronim)
    df = df[df[geo_col].isin(EU27)].copy()

    if df.empty:
        raise ValueError(
            f"dupa filtrare raman 0 randuri pentru EU27 (filtre: {filters}). "
            f"Cel putin doua filtre sunt reciproc incompatibile (nu co-exista simultan "
            f"in date) — nu e vorba de o valoare inexistenta pe o dimensiune izolata "
            f"(altfel apare avertismentul [!] obisnuit), ci de o combinatie contradictorie."
        )

    # daca dupa filtrare raman mai multe randuri per tara -> filtre insuficiente
    dubluri = df[geo_col].duplicated().sum()
    if dubluri > 0 and prefer is not None:
        mask = pd.Series(False, index=df.index)
        for c in dim_cols:
            mask = mask | (df[c].astype(str) == str(prefer))
        if mask.any():
            df = df[mask]
            print(f"   ({acronim}: am pastrat randurile cu '{prefer}'.)")
        dubluri = df[geo_col].duplicated().sum()
    if dubluri > 0:
        print(f"   [!] {acronim}: {dubluri} randuri duplicate per tara dupa filtrare "
              f"(filtre insuficiente) — pastrez prima aparitie.")
        df = df.drop_duplicates(subset=[geo_col], keep="first")

    df = df.set_index(geo_col)

    # coloanele-an existente in dataset (ca int), apoi reindexam la 2014-2024
    year_map = {int(str(c).strip()): c for c in year_cols}
    tabel = pd.DataFrame(index=EU27, columns=ANI, dtype="float")
    for an in ANI:
        if an in year_map and year_map[an] in df.columns:
            tabel[an] = df[year_map[an]].reindex(EU27)
        # altfel ramane NaN (anul nu exista in dataset)

    # garda: foaie complet goala (randuri prezente dar toate valorile NaN pe 2014-2024)
    if tabel.isna().all().all():
        raise ValueError(
            f"tabel COMPLET gol pe {ANI[0]}-{ANI[-1]} desi filtrarea a lasat randuri "
            f"(filtre: {filters}). Filtrul nimereste o serie fara valori in acesti ani. "
            f"Dimensiuni reale disponibile in {code}: {inventar}"
        )
    return tabel


def build_sheet(tabel):
    """Adauga cod_tara + nume tara in fata tabelei geo x ani."""
    out = tabel.reindex(EU27).reset_index()
    out = out.rename(columns={"index": "cod_tara"})
    out.insert(1, "tara", out["cod_tara"].map(NUME_TARI))
    out.columns = ["cod_tara", "tara"] + [str(a) for a in ANI]
    return out


# ------------------------------------------------------------------------------
# CONSTRUCTII (variabile derivate)
# ------------------------------------------------------------------------------
def construieste_urb_share(acronim="Urb_share"):
    """
    Ponderea populatiei URBANE din urt_pjanaggr3, pe baza tipologiei urban-rural.
    STRUCTURA CONFIRMATA (Eurostat Glossary: Urban-rural typology):
      terrtypo partitioneaza populatia in 3 categorii care insumate dau totalul:
        URB = predominant urban, INT = intermediar, RUR = predominant rural.
      Restul codurilor (CST_R, MNT_R, ISL_R, BRD_R, INT_CTC, RUR_RMT etc.) sunt
      clasificari SEPARATE / suprapuse -> NU se aduna cu URB/INT/RUR.
    Construim: Urb_share = URB / (URB + INT + RUR) * 100, cu sex=T, age=TOTAL.
    Nota: CY si MT nu au date in acest dataset (state fara diferentiere NUTS3
    in tipologie) -> raman NaN si le decizi manual (ambele ~100% urban de facto).
    """
    code = "urt_pjanaggr3"
    print(f"\n>> {acronim} ({code}) — URB / (URB+INT+RUR)")
    try:
        df = eurostat.get_data_df(code)
    except Exception as e:
        print(f"   [!] nu am putut descarca {code}: {e} — sar peste Urb_share.")
        return None

    geo_col = find_geo_col(df)
    year_cols = find_year_cols(df)
    dim_cols = dimension_cols(df, geo_col, year_cols)

    # gaseste dimensiunea tipologiei (cea care contine URB/INT/RUR)
    dim_terit = None
    for c in dim_cols:
        vals = set(map(str, df[c].dropna().unique()))
        if {"URB", "INT", "RUR"}.issubset(vals):
            dim_terit = c
            break
    if dim_terit is None:
        print("   [!] Nu gasesc dimensiunea cu URB/INT/RUR. Dimensiuni disponibile:")
        for c in dim_cols:
            vals = sorted(map(str, df[c].dropna().unique()))
            print(f"      - {str(c).split(chr(92))[0]}: {vals[:20]}")
        return None

    # filtreaza sex=T si age=TOTAL daca dimensiunile exista (ca sa nu dublam)
    base = df.copy()
    for dimname, val in (("sex", "T"), ("age", "TOTAL")):
        col = next((c for c in dim_cols if str(c).split("\\")[0].lower() == dimname), None)
        if col is not None and val in set(base[col].astype(str)):
            base = base[base[col].astype(str) == val]

    year_map = {int(str(c).strip()): c for c in year_cols}

    def categorie(val):
        sub = base[base[dim_terit].astype(str) == val]
        sub = sub[sub[geo_col].isin(EU27)].drop_duplicates(subset=[geo_col], keep="first")
        sub = sub.set_index(geo_col)
        tab = pd.DataFrame(index=EU27, columns=ANI, dtype="float")
        for an in ANI:
            if an in year_map and year_map[an] in sub.columns:
                tab[an] = sub[year_map[an]].reindex(EU27)
        return tab

    u = categorie("URB")
    i = categorie("INT")
    r = categorie("RUR")
    # o categorie lipsa (tara fara acel tip de regiune) inseamna 0, nu 'necunoscut';
    # dar daca nu exista date DELOC (total=0), lasam NaN.
    total = u.fillna(0) + i.fillna(0) + r.fillna(0)
    share = (u.fillna(0) / total * 100).round(2)
    share = share.where(total > 0)

    n_lipsa = int(share.isna().sum().sum())
    tari_goale = [t for t in EU27 if share.loc[t].isna().all()]
    print(f"   URB/(URB+INT+RUR) calculat. Celule lipsa: {n_lipsa}. "
          f"Tari fara date: {tari_goale if tari_goale else 'niciuna'}.")
    print(f"   ATENTIE: tipologia clasifica regiuni NUTS3 intregi -> tarile mono-regionale "
          f"(LU/SI/CY/MT si posibil altele mici) pot iesi 0% sau 100%, nerealist.")
    return share


def _band_age_codes(age_values, lo, hi):
    """Alege codurile de varsta care compun banda [lo, hi] (hi=None => deschis sus).
    Foloseste DOAR ani individuali Y<n> + codul deschis Y_GE<n>, ca sa nu dubleze
    prin intervale agregate. Pentru banda deschisa, daca exista Y_GE(k) cu k<=lo,
    il foloseste direct (ex. Y_GE15 pt '15+')."""
    individ = {}       # n -> cod Y<n>
    open_top = None    # (n, cod) pt cel mai mic Y_GE<n>
    for c in age_values:
        s = str(c)
        m = re.fullmatch(r"Y(\d+)", s)
        if m:
            individ[int(m.group(1))] = c
            continue
        m = re.fullmatch(r"Y_GE(\d+)", s)
        if m:
            n = int(m.group(1))
            if open_top is None or n < open_top[0]:
                open_top = (n, c)
    use_open = (hi is None and open_top is not None)
    chosen = []
    for n, c in sorted(individ.items()):
        if n < lo:
            continue
        if hi is not None and n > hi:
            continue
        if use_open and n >= open_top[0]:   # acoperit de codul deschis
            continue
        chosen.append(c)
    if use_open:
        chosen.append(open_top[1])
    return chosen


def construieste_pjanmarsta(marriages_raw):
    """Construieste din demo_pjanmarsta:
       - Married_share_15p    = MAR / TOTAL, varsta 15+
       - Married_share_25_49  = MAR / TOTAL, varsta 25-49
       - Mar_rate_unmarried   = Marriages / (SIN+WID+DIV) 15+ * 1000
    Returneaza (dict foi geo x ani, list info_rows)."""
    code = "demo_pjanmarsta"
    print(f"\n>> demo_pjanmarsta — ponderi casatoriti + rata pe populatia necasatorita")
    try:
        df = eurostat.get_data_df(code)
    except Exception as e:
        print(f"   [!] nu am putut descarca {code}: {e} — sar peste aceste variabile.")
        return {}, []

    geo_col = find_geo_col(df)
    year_cols = find_year_cols(df)
    dim_cols = dimension_cols(df, geo_col, year_cols)
    year_map = {int(str(c).strip()): c for c in year_cols}

    def col(name):
        return next((c for c in dim_cols if str(c).split("\\")[0].lower() == name), None)
    sex_c, marsta_c, age_c = col("sex"), col("marsta"), col("age")
    if None in (sex_c, marsta_c, age_c):
        print(f"   [!] lipsesc dimensiuni asteptate (sex/marsta/age). Am: {[str(c) for c in dim_cols]}")
        return {}, []

    age_values = sorted(map(str, df[age_c].dropna().unique()))
    codes_15p = _band_age_codes(age_values, 15, None)
    codes_2549 = _band_age_codes(age_values, 25, 49)
    print(f"   Coduri varsta 15+   ({len(codes_15p)}): {codes_15p[:8]}{'...' if len(codes_15p)>8 else ''}")
    print(f"   Coduri varsta 25-49 ({len(codes_2549)}): {codes_2549}")
    if not codes_15p or not codes_2549:
        print("   [!] Nu am putut compune benzile de varsta din codurile disponibile:")
        print(f"       age: {age_values}")
        return {}, []

    def band_pop(marsta_vals, age_codes, sex="T"):
        sub = df[(df[sex_c].astype(str) == sex)
                 & (df[marsta_c].astype(str).isin(marsta_vals))
                 & (df[age_c].astype(str).isin(age_codes))]
        sub = sub[sub[geo_col].isin(EU27)]
        tab = pd.DataFrame(index=EU27, columns=ANI, dtype="float")
        for an in ANI:
            if an in year_map and year_map[an] in sub.columns:
                g = sub.groupby(geo_col)[year_map[an]].sum(min_count=1)
                tab[an] = g.reindex(EU27)
        return tab

    mar15 = band_pop(["MAR"], codes_15p)
    tot15 = band_pop(["TOTAL"], codes_15p)
    mar2549 = band_pop(["MAR"], codes_2549)
    tot2549 = band_pop(["TOTAL"], codes_2549)
    unm15 = band_pop(["SIN", "WID", "DIV"], codes_15p)

    foi_new = {}
    info_new = []

    share15 = (mar15 / tot15 * 100).round(2)
    foi_new["Married_share_15p"] = build_sheet(share15)
    info_new.append(["Married_share_15p", "Pondere casatoriti in pop. 15+", code,
                     "MAR/TOTAL, age>=15, sex=T", "derivat",
                     "Candidat dependenta. Necasatorit vs total pe varsta matrimoniala."])

    share2549 = (mar2549 / tot2549 * 100).round(2)
    foi_new["Married_share_25_49"] = build_sheet(share2549)
    info_new.append(["Married_share_25_49", "Pondere casatoriti in pop. 25-49", code,
                     "MAR/TOTAL, age 25-49, sex=T", "derivat",
                     "Candidat dependenta, varste de formare a familiei (mai discriminant)."])

    if marriages_raw is not None:
        rate = (marriages_raw / unm15 * 1000).round(3)
        foi_new["Mar_rate_unmarried"] = build_sheet(rate)
        info_new.append(["Mar_rate_unmarried", "Rata nuptialitatii pop. necasatorite 15+", code + " + demo_nind",
                         "MARRIAGE / (SIN+WID+DIV, 15+) * 1000", "derivat",
                         "Rata generala de nuptialitate (numitor = pop. expusa, nu totala)."])
    else:
        print("   [!] Marriages_raw indisponibil -> nu construiesc Mar_rate_unmarried.")

    for nume, tab in [("Married_share_15p", share15), ("Married_share_25_49", share2549)]:
        n_lipsa = int(tab.isna().sum().sum())
        print(f"   {nume}: {n_lipsa} celule lipsa. Ex. RO 2014 = {tab.loc['RO', 2014]}")
    return foi_new, info_new


# ------------------------------------------------------------------------------
# RULARE PRINCIPALA
# ------------------------------------------------------------------------------
def main():
    print("Descarc si asamblez baza de date 2014-2024 (EU27, nivel de tara)...\n")

    foi = {}          # nume_foaie -> DataFrame (formatat pt Excel)
    tabele = {}       # acronim -> tabela geo x ani (pt derivate)
    info_rows = []

    # 1) Variabilele directe
    for acronim, (code, filters, descr) in DIRECTE.items():
        print(f">> {acronim} ({code})")
        try:
            tab = fetch_series(code, filters, acronim)
            tabele[acronim] = tab
            foi[acronim] = build_sheet(tab)
            n_lipsa = int(tab.isna().sum().sum())
            print(f"   OK — {n_lipsa} celule lipsa din {27*len(ANI)}.")
            info_rows.append([acronim, descr, code,
                              "; ".join(f"{k}={v}" for k, v in filters.items()) or "-",
                              "direct", NOTE.get(acronim, "")])
        except Exception as e:
            print(f"   [EROARE] {acronim}: {e}")
            info_rows.append([acronim, descr, code, "-", "EROARE", str(e)])

    # 2) Gen_gap: indicator DEDICAT la nivel de tara (sdg_05_30), nu construit.
    #    tepsr_lm220 din tabelul initial era versiunea regionala NUTS-2, oprita in 2019.
    print("\n>> Gen_gap (sdg_05_30) — indicator dedicat, ocupare totala (EMP)")
    try:
        gg = fetch_series("sdg_05_30",
                          {"age": "Y20-64", "unit": "PC_PNT", "wstatus": "EMP"},
                          "Gen_gap", prefer="EMP")
        tabele["Gen_gap"] = gg
        foi["Gen_gap"] = build_sheet(gg)
        n_lipsa = int(gg.isna().sum().sum())
        print(f"   OK — {n_lipsa} celule lipsa din {27*len(ANI)}.")
        info_rows.append(["Gen_gap", "Diferenta de gen in ocupare 20-64 (M-F)", "sdg_05_30",
                          "age=Y20-64; unit=PC_PNT; wstatus=EMP", "direct",
                          "Inlocuieste tepsr_lm220 (regional NUTS2, oprit 2019). Serie 2009-2024."])
    except Exception as e:
        print(f"   [EROARE] Gen_gap: {e}")
        info_rows.append(["Gen_gap", "Diferenta de gen in ocupare 20-64 (M-F)", "sdg_05_30",
                          "-", "EROARE", str(e)])

    # 3) EmpF_x_EmpM: interactiune — nu are echivalent direct, ramane construita
    if "Emp_rate_F" in tabele and "Emp_rate_M" in tabele:
        emp_prod = (tabele["Emp_rate_F"] * tabele["Emp_rate_M"]).round(2)
        foi["EmpF_x_EmpM"] = build_sheet(emp_prod)
        info_rows.append(["EmpF_x_EmpM", "Interactiune ocupare femei x barbati", "derivat",
                          "Emp_rate_F * Emp_rate_M", "derivat",
                          "Produs brut de procente — scalarea/log e decizie de Pasul 3."])
        print(">> EmpF_x_EmpM — construit din Emp_rate_F * Emp_rate_M. OK")

    # 3) Urb_share (best-effort)
    urb = construieste_urb_share()
    if urb is not None:
        foi["Urb_share"] = build_sheet(urb)
        n_lipsa = int(urb.isna().sum().sum())
        info_rows.append(["Urb_share", "Pondere populatie urbana (urban/total)", "urt_pjanaggr3",
                          "urban/total", "derivat",
                          f"CY si MT probabil lipsa (state-insula). {n_lipsa} celule lipsa."])
        print(f"   OK — {n_lipsa} celule lipsa.")

    # 4) demo_pjanmarsta: ponderi casatoriti (15+ si 25-49) + rata pe necasatoriti 15+
    foi_pjan, info_pjan = construieste_pjanmarsta(tabele.get("Marriages_raw"))
    foi.update(foi_pjan)
    info_rows.extend(info_pjan)

    # 5) Foaia Info
    info = pd.DataFrame(info_rows, columns=["variabila", "descriere", "cod_dataset",
                                            "filtre", "tip", "note"])

    # ordinea foilor: Info + grupat logic
    ordine = ["Mar_rate", "Mar_rate_prim_F", "Mar_rate_prim_M",
              "Age_first_mar_F", "Age_first_mar_M", "Marriages_raw",
              "Married_share_15p", "Married_share_25_49", "Mar_rate_unmarried",
              "GDP_cap", "Une_rate", "Emp_rate", "Emp_rate_F", "Emp_rate_M",
              "Gen_gap", "EmpF_x_EmpM",
              "Ter_educ", "NEET", "Tot_fer_r", "Age_leave",
              "Pov_risk", "Mat_depriv", "Subj_pov",
              "Urb_share", "Net_migr_rate", "Liv_parent"]

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as xl:
        info.to_excel(xl, sheet_name="Info", index=False)
        for acronim in ordine:
            if acronim in foi:
                foi[acronim].to_excel(xl, sheet_name=acronim[:31], index=False)

    print(f"\n[OK] Scris: {OUTPUT_XLSX} — {len(foi)} foi de date + Info.")
    print("     (Excluse din acest fisier, le facem la Pasul 3 dupa ce decizi breakdown-urile:")
    print("      Pop_marital_status/ponderea casatoritilor pe grupe de varsta, Mar_rate_unmarried, Liv_parent_dummy.)")


if __name__ == "__main__":
    main()
