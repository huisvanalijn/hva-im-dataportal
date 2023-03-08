import pandas as pd
import datetime

#when deploying change to: df_collectie = pd.read_csv(r'/home/floreverkest/hva-im-dataportal/static/data/hva/collectie.csv', delimiter=';', low_memory=False)
df_collectie = pd.read_csv(r'dataportalapp\static\data\hva\collectie.csv', delimiter=';', low_memory=False)

df_thesaurus = pd.read_csv(r'C:\Users\Verkesfl\OneDrive - Groep Gent\Bureaublad\hva-im-dataportal\dataportalapp\static\data\hva\thesaurus.csv', delimiter=';', low_memory=False)
df_associaties = pd.read_excel(r'C:\Users\Verkesfl\OneDrive - Groep Gent\Bureaublad\hva-im-dataportal\dataportalapp\static\data\hva\overzichtassociaties.xlsx')

df_rschijf = pd.read_excel(r'C:\Users\Verkesfl\OneDrive - Groep Gent\Bureaublad\hva-im-dataportal\dataportalapp\static\data\hva\rschijf.xlsx')
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains("WERKMAP", na=False)]
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains("EXTERNE SCHIJVEN", na=False)]
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains(r"_", na=False)]
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains(r"A3", na=False)]
df_rschijflimited = df_rschijf[(df_rschijf["objectnummer"].str.startswith("FO-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("DB-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("F0", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("AU-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("19", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("20", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("DIA-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("AF", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("RE-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("VI-", na=False))]
df_rschijflimited = df_rschijflimited[~df_rschijflimited['objectnummer'].str.contains(r"_", na=False)]
df_rschijflimited = df_rschijflimited[~((df_rschijflimited['objectnummer'].str.contains("a", na=False))
                    | (df_rschijf['objectnummer'].str.contains("b", na=False))
                    | (df_rschijf['objectnummer'].str.contains("kopie", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"\(", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r" 2", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"\)", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"c", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"d", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r" 3", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r" 4", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"C", na=False)))]

#if columnnames change (due to different CMS), change the names here:
df_collectie['instelling.naam'] = df_collectie['instelling.naam']
df_collectie['instelling.code'] = df_collectie['instelling.code']
df_collectie['afdeling'] = df_collectie['afdeling']
df_collectie['onderscheidende_kenmerken'] = df_collectie['onderscheidende_kenmerken']
df_collectie['objectnummer'] = df_collectie['objectnummer']
df_collectie['objectnaam'] = df_collectie['objectnaam']
df_collectie['titel'] = df_collectie['titel']
df_collectie['reproductie.referentie'] = df_collectie['reproductie.referentie']
df_collectie['inhoud.onderwerp'] = df_collectie['inhoud.onderwerp']
df_collectie['inhoud.persoon.naam'] = df_collectie['inhoud.persoon.naam']
df_collectie['associatie.onderwerp'] = df_collectie['associatie.onderwerp']
df_collectie['associatie.periode'] = df_collectie['associatie.periode']
df_collectie['associatie.persoon'] = df_collectie['associatie.persoon']
df_collectie['vervaardiger'] = df_collectie['vervaardiger']
df_collectie['vervaardiger.rol'] = df_collectie['vervaardiger.rol']
df_collectie['vervaardiging.datum.begin'] = df_collectie['vervaardiging.datum.begin']
df_collectie['vervaardiging.datum.begin.prec'] = df_collectie['vervaardiging.datum.begin.prec']
df_collectie['vervaardiging.datum.eind'] = df_collectie['vervaardiging.datum.eind']
df_collectie['vervaardiging.datum.eind.prec'] = df_collectie['vervaardiging.datum.eind.prec']
df_collectie['vervaardiging.periode'] = df_collectie['vervaardiging.periode']
df_collectie['vervaardiging.plaats'] = df_collectie['vervaardiging.plaats']
df_collectie['afmeting.eenheid'] = df_collectie['afmeting.eenheid']
df_collectie['afmeting.waarde'] = df_collectie['afmeting.waarde']
df_collectie['afmeting.soort'] = df_collectie['afmeting.soort']
df_collectie['rechten.type'] = df_collectie['rechten.type']
df_collectie['rechten.startdatum'] = df_collectie['rechten.startdatum']
df_collectie['rechten.machtigingsstatus'] = df_collectie['rechten.machtigingsstatus']
df_collectie['rechten.uit.aanvrager'] = df_collectie['rechten.uit.aanvrager']
df_collectie['rechten.referentienummer'] = df_collectie['rechten.referentienummer']
df_collectie['rechten.bijzonderheden'] = df_collectie['rechten.bijzonderheden']
df_collectie['verwerving.bron'] = df_collectie['verwerving.bron']
df_collectie['verwerving.datum'] = df_collectie['verwerving.datum']
df_collectie['verwerving.methode'] = df_collectie['verwerving.methode']
df_collectie['huidige_standplaats'] = df_collectie['huidige_standplaats']
df_collectie['standplaats.standaard'] = df_collectie['standplaats.standaard']
df_collectie['huidige_standplaats.context'] = df_collectie['huidige_standplaats.context']
df_collectie['toestand'] = df_collectie['toestand']
df_collectie['webpublicatie'] = df_collectie['webpublicatie']
df_collectie['wijziging.datum'] = df_collectie['wijziging.datum']
df_collectie['wijziging.naam'] = df_collectie['wijziging.naam']

year = datetime.datetime.now().year

##################################################################################################################################################
# 1. COLLECTIE
##################################################################################################################################################

# foutieve instellingsnaam
def h001():
    df_001 = df_collectie[df_collectie["instelling.naam"] != 'Het Huis van Alijn (Gent)']
    return df_001

#foutieve instellingscode
def h002():
    df_002 = df_collectie[df_collectie["instelling.code"] != 'INST-570']
    return df_002

# objectnummer foutieve start
def h003():
    df_003 = df_collectie[~df_collectie['objectnummer'].str.startswith(('AU-', 'FO-', '19', '20', 'DIA-', 'AF', 'DB-', 'RE-', 'F0', 'VI'))]
    return df_003

# objectnummer foutieve format:
def h004():
    df_01 = df_collectie[df_collectie['objectnummer'].str.startswith('AU-')]
    df_01 = df_01[~df_01['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    df_02 = df_collectie[df_collectie['objectnummer'].str.startswith('DIA-')]
    df_02 = df_02[~df_02['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    df_02 = df_02[~df_02['objectnummer'].apply(lambda x: len(str(x)) == 13)]
    df_03 = df_collectie[df_collectie['objectnummer'].str.startswith('FO-')]
    df_03 = df_03[~df_03['objectnummer'].apply(lambda x: len(str(x)) == 11)]
    df_03 = df_03[~df_03['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    df_04 = df_collectie[df_collectie['objectnummer'].str.startswith('RE-')]
    df_04 = df_04[~df_04['objectnummer'].apply(lambda x: len(str(x)) == 10)]
    df_05 = df_collectie[df_collectie['objectnummer'].str.startswith('F0-')]
    df_05 = df_05[~df_05['objectnummer'].apply(lambda x: len(str(x)) == 6)]
    df_06 = df_collectie[df_collectie['objectnummer'].str.startswith('VI-')]
    df_06 = df_06[~df_06['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    df_07 = df_collectie[df_collectie['objectnummer'].str.startswith('AF-')]
    df_07 = df_07[~df_07['objectnummer'].apply(lambda x: len(str(x)) == 11)]
    df_08 = df_collectie[df_collectie['objectnummer'].str.startswith('19-')]
    df_08 = df_08[~df_08['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    # df_08 = df_08[~df_08['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    df_09 = df_collectie[df_collectie['objectnummer'].str.startswith('20-')]
    df_09 = df_09[~df_09['objectnummer'].apply(lambda x: len(str(x)) == 12)]
    # df_09 = df_09[~df_09['objectnummer'].apply(lambda x: len(str(x)) == 8)]
    df_10 = df_collectie[df_collectie['objectnummer'].str.startswith('DB-')]
    df_10 = df_10[~df_10['objectnummer'].apply(lambda x: len(str(x)) == 15)]
    frames = [df_01, df_02, df_03, df_04, df_05, df_06, df_07, df_08, df_09, df_10]
    df_004 = pd.concat(frames)
    return df_004

# foutief onderscheidend kenmerk
def h005():
    df_005 = df_collectie[df_collectie['onderscheidende_kenmerken'] != 'DIGITALE COLLECTIE']
    df_005 = df_005[df_005['onderscheidende_kenmerken'] != 'OBJECT']
    df_005 = df_005[df_005['onderscheidende_kenmerken'] != 'BEELD']
    df_005 = df_005[df_005['onderscheidende_kenmerken'] != 'DOCUMENTAIRE COLLECTIE']
    return df_005

# ontbrekende basisregistratie (objectnaam, titel, afbeelding, associatie)
def h006():
    df_006 = df_collectie[df_collectie['objectnaam'].isna()]
    return df_006

def h007():
    df_007 = df_collectie[df_collectie['titel'].isna()]
    return df_007

def h008():
    df_008 = df_collectie[df_collectie['reproductie.referentie'].isna()]
    df_008 = pd.merge(df_008, df_rschijflimited, on="objectnummer", how="outer")
    df_008 = df_008[~df_008['instelling.naam'].isna()]
    return df_008

def h009():
    df_009 = df_collectie[df_collectie['associatie.onderwerp'].isna()]
    return df_009

# ERROR NOT WORKING: vervaardiging plaats niet in associatie onderwerp
def h010():
    df_010 = df_collectie[df_collectie['vervaardiging.plaats'].notna()]
    df_010 = df_010[df_010['vervaardiging.plaats'] != '$']
    df_010['vervaardiging.plaats'] = df_010['vervaardiging.plaats'].map(lambda x: x.lstrip('$').rstrip('$'))
    df_010 = df_010[~df_010["vervaardiging.plaats"].isin(df_010["associatie.onderwerp"])]
    return df_010

# vervaardiging datering niet in associatie periode
def h011():
    df_011 = df_collectie[df_collectie['vervaardiging.datum.begin'].notna()]
    df_011 = df_011[df_011['associatie.periode'].isna()]
    return df_011

# vervaardiging datum begin = vervaardiging datum eind
def h012():
    df_012 = df_collectie.loc[df_collectie['vervaardiging.datum.begin'] == df_collectie['vervaardiging.datum.eind']]
    return df_012

# vervaardiging datum begin > vervaardiging datum eind
def h013():
    df_013 = df_collectie.loc[df_collectie['vervaardiging.datum.begin'] > df_collectie['vervaardiging.datum.eind']]
    return df_013

# datering niet in correcte formaat
def h014():
    df_014 = df_collectie[df_collectie['vervaardiging.datum.begin'].notna()]
    df_014 = df_014.drop(df_014[pd.to_datetime(df_014['vervaardiging.datum.begin'], format='%Y-%m',
                                            errors='coerce').notna()].index)
    df_014 = df_014[~df_014['vervaardiging.datum.begin'].str.startswith('14')]
    df_014 = df_014[~df_014['vervaardiging.datum.begin'].str.startswith('15')]
    df_014 = df_014[~df_014['vervaardiging.datum.begin'].str.startswith('16')]
    return df_014

# ontbrekende afmetingen
def h015():
    df_015 = df_collectie[df_collectie['afmeting.waarde'].isna()]
    return df_015

# afmetingen objecten niet in cm
def h016():
    df_016 = df_collectie[df_collectie['onderscheidende_kenmerken'] == 'OBJECT']
    df_016 = df_016[~df_016['afmeting.eenheid'].isna()]
    df_016 = df_016[~df_016['afmeting.eenheid'].str.startswith('cm')]
    return df_016

# afmetingen DB niet in min, kb, mb of gb
def h017():
    df_017 = df_collectie[df_collectie['onderscheidende_kenmerken'] == 'DIGITALE COLLECTIE']
    df_017 = df_017[~df_017['afmeting.eenheid'].isna()]
    df_017 = df_017[~df_017['afmeting.eenheid'].str.startswith('min')]
    df_017 = df_017[~df_017['afmeting.eenheid'].str.startswith('kB')]
    df_017 = df_017[~df_017['afmeting.eenheid'].str.startswith('MB')]
    df_017 = df_017[~df_017['afmeting.eenheid'].str.startswith('GB')]
    return df_017

# afmetingen documenten niet in mm
def h018():
    df_018 = df_collectie[df_collectie['onderscheidende_kenmerken'] == 'DOCUMENTAIRE COLLECTIE|BEELD']
    df_018 = df_018[~df_018['afmeting.eenheid'].isna()]
    df_018 = df_018[~df_018['afmeting.eenheid'].str.startswith('mm')]
    return df_018

# ontbrekende rechtenstatus
def h019():
    df_019 = df_collectie[df_collectie['rechten.type'].isna()]
    return df_019

# ontbrekende machtigingsstatus
def h020():
    df_020 = df_collectie[~df_collectie['rechten.type'].isna()]
    df_020 = df_020[~df_020['rechten.machtigingsstatus'].str.contains("toegewezen", na=False)]
    return df_020

# ontbrekend formulier
def h021():
    df_021 = df_collectie[(df_collectie['rechten.type'] == "IN COPYRIGHT - NON-COMMERCIAL USE PERMITTED") |
                    (df_collectie['rechten.type'] == "CC-BY-NC 4.0") |
                    (df_collectie['rechten.type'] == "CC-BY-SA 4.0")]
    df_021 = df_021[df_021['rechten.referentienummer'].isna()]
    return df_021

# publiek domein OP BASIS VAN STERFTEDATUM VERVAARDIGER NOG TOEVOEGEN!
def h022():
    df_022 = df_collectie[df_collectie['rechten.type'] != 'PUBLIC DOMAIN']
    df_022_1 = df_022[df_022['associatie.periode'].str.contains("18de eeuw", na=False)]
    df_022_1 = df_022_1[~df_022_1['associatie.periode'].str.contains("19de eeuw", na=False)]

    df_022_2 = df_022[df_022['vervaardiging.datum.eind'].isna()]
    df_022_2 = df_022_2[~df_022_2['vervaardiging.datum.begin'].isna()]
    df_022_2['vervaardiging.datum.begin'] = df_022_2['vervaardiging.datum.begin'].astype(str)
    df_022_2['vervaardiging.datum.begin'] = df_022_2['vervaardiging.datum.begin'].str[:4]
    df_022_2['vervaardiging.datum.begin'] = df_022_2['vervaardiging.datum.begin'].astype(int)
    df_022_2 = df_022_2[df_022_2['vervaardiging.datum.begin'] <= (year - 150)]

    df_022_3 = df_022[~df_022['vervaardiging.datum.eind'].isna()]
    df_022_3['vervaardiging.datum.eind'] = df_022_3['vervaardiging.datum.eind'].astype(str)
    df_022_3['vervaardiging.datum.eind'] = df_022_3['vervaardiging.datum.eind'].str[:4]
    df_022_3 = df_022_3[df_022_3['vervaardiging.datum.eind'] != '$']
    df_022_3['vervaardiging.datum.eind'] = df_022_3['vervaardiging.datum.eind'].astype(int)
    df_022_3 = df_022_3[df_022_3['vervaardiging.datum.eind'] <= (year - 150)]

    frames = [df_022_1, df_022_2, df_022_3]
    df_022 = pd.concat(frames)
    return df_022

# ontbrekende/foutieve toestand
def h023():
    df_023 = df_collectie[df_collectie['onderscheidende_kenmerken'] != 'DIGITALE COLLECTIE']
    df_023 = df_collectie[~df_collectie['toestand'].str.contains("goed", na=False)]
    df_023 = df_023[~df_023['toestand'].str.contains("matig", na=False)]
    df_023 = df_023[~df_023['toestand'].str.contains("slecht", na=False)]
    return df_023

# ontbrekende/foutieve verwervingsmethode
def h024():
    df_024 = df_collectie[df_collectie['verwerving.methode'] != 'schenking']
    df_024 = df_024[df_024['verwerving.methode'] != 'aankoop']
    df_024 = df_024[df_024['verwerving.methode'] != 'onbekend']
    df_024 = df_024[df_024['verwerving.methode'] != 'bruikleen']
    return df_024

# Gent in titel maar niet in associatie
def h025():
    df_025 = df_collectie[df_collectie["titel"].str.contains("Gent", na=False)]
    df_025 = df_025[~df_025["associatie.onderwerp"].str.contains("Gent", na=False)]
    return df_025

# wereldtentoonstelling X in associatie maar niet stad apart als associatie
def h026():
    df_026_01 = df_collectie[df_collectie["associatie.onderwerp"].str.contains("Wereldtentoonstelling Brussel", na=False)]
    search = ['$Brussel', 'Brussel$']
    df_026_01 = df_026_01[~df_026_01["associatie.onderwerp"].str.contains('|'.join(search), na=False)]
    df_026_02 = df_collectie[df_collectie["associatie.onderwerp"].str.contains("Wereldtentoonstelling Antwerpen", na=False)]
    search = ['$Antwerpen', 'Antwerpen$']
    df_026_02 = df_026_02[~df_026_02["associatie.onderwerp"].str.contains('|'.join(search), na=False)]
    df_026_03 = df_collectie[df_collectie["associatie.onderwerp"].str.contains("Wereldtentoonstelling Gent", na=False)]
    search = ['$Gent', 'Gent$']
    df_026_03 = df_026_03[~df_026_03["associatie.onderwerp"].str.contains('|'.join(search), na=False)]
    df_026_04 = df_collectie[df_collectie["associatie.onderwerp"].str.contains("Wereldtentoonstelling Luik", na=False)]
    search = ['$Luik', 'Luik$']
    df_026_04 = df_026_04[~df_026_04["associatie.onderwerp"].str.contains('|'.join(search), na=False)]
    df_026_05 = df_collectie[df_collectie["associatie.onderwerp"].str.contains("Wereldtentoonstelling Parijs", na=False)]
    search = ['$Parijs', 'Parijs$']
    df_026_05 = df_026_05[~df_026_05["associatie.onderwerp"].str.contains('|'.join(search), na=False)]
    df_026_06 = df_collectie[df_collectie["associatie.onderwerp"].str.contains("Wereldtentoonstelling Londen", na=False)]
    search = ['$Londen', 'Londen$']
    df_026_06 = df_026_06[~df_026_06["associatie.onderwerp"].str.contains('|'.join(search), na=False)]
    df_026_07 = df_collectie[df_collectie["associatie.onderwerp"].str.contains("Wereldtentoonstelling Amsterdam", na=False)]
    search = ['$Amsterdam', 'Amsterdam$']
    df_026_07 = df_026_07[~df_026_07["associatie.onderwerp"].str.contains('|'.join(search), na=False)]
    frames = [df_026_01, df_026_02, df_026_03, df_026_04, df_026_05, df_026_06, df_026_07]
    df_026 = pd.concat(frames)
    return df_026

# titel eindigt op .
def h027():
    df_027 = df_collectie[df_collectie['titel'].str.endswith(".", na=False)]
    return df_027

# titel begint met hoofdletter
def h028():
    df_028 = df_collectie[~df_collectie['titel'].isna()]
    df_028['t'] = df_028['titel'].astype(str).str[0]
    df_028 = df_028[~df_028['t'].str.isupper()]
    df_028 = df_028[~df_028['t'].str.isdigit()]
    df_028 = df_028[~df_028['titel'].str.startswith("'s ", na=False)]
    df_028 = df_028[~df_028['titel'].str.startswith("'t ", na=False)]
    return df_028

# titel bevat (, ) of "
def h029():
    search = [r'\)', r'\(', '"']
    df_029 = df_collectie[df_collectie['titel'].str.contains('|'.join(search), na=False)]
    return df_029

# associatie komt niet voor in lijst associaties
def h030():
    df_asso = df_collectie['associatie.onderwerp'].str.split('$', expand=True)
    aantal = len(df_asso.columns)

    xs = []
    for i in range(aantal):
        xs.append(i)

    df_030 = pd.concat([df_asso[xs].melt(value_name='associaties')])
    df_030.mask(df_030.eq('None')).dropna()
    df_030 = df_030[df_030['associaties'].notna()]
    df_030 = df_030['associaties'].drop_duplicates()
    df_030 = df_030.to_frame()
    df_plaatsen = df_thesaurus[((df_thesaurus['term.soort'].str.contains("plaats", na=False)) | (df_thesaurus['term.soort'].str.contains("geografisch", na=False)))]
    df_030["exists"] = df_030['associaties'].isin(df_plaatsen["term"])
    df_030 = df_030.mask(df_030.eq(True)).dropna()
    df_030['present'] = df_030['associaties'].isin(df_associaties['Trefwoord'])
    df_030 = df_030.mask(df_030.eq(True)).dropna()
    return df_030

# foutieve vervaardiging datum begin precisie
def h031():
    df_031 = df_collectie[df_collectie['vervaardiging.datum.begin.prec'].notna()]
    df_031 = df_031[df_031['vervaardiging.datum.begin.prec'] != 'ca.']
    df_031 = df_031[df_031['vervaardiging.datum.begin.prec'] != '$']
    df_031 = df_031[df_031['vervaardiging.datum.begin.prec'] != 'na']
    df_031 = df_031[df_031['vervaardiging.datum.begin.prec'] != 'voor']
    df_031 = df_031[df_031['vervaardiging.datum.begin.prec'] != 'vanaf']
    return df_031

# foutieve vervaardiging datum eind precisie
def h032():
    df_032 = df_collectie[df_collectie['vervaardiging.datum.eind.prec'].notna()]
    df_032 = df_032[df_032['vervaardiging.datum.eind.prec'] != 'ca.']
    df_032 = df_032[df_032['vervaardiging.datum.eind.prec'] != '$']
    df_032 = df_032[df_032['vervaardiging.datum.eind.prec'] != 'voor']
    return df_032

# ontbrekende uitleg bij bijzonderheden bij records in PD
def h033():
    df_033 = df_collectie[df_collectie['rechten.type'] == 'PUBLIC DOMAIN']
    df_033 = df_033[df_033['rechten.bijzonderheden'].isna()]
    return df_033

##################################################################################################################################################
# 2. THESAURUS
##################################################################################################################################################

# bron afwezig of niet correct
def ht01():
    df_t01 = df_thesaurus[df_thesaurus['bron'] != 'http://vocab.getty.edu/aat/']
    df_t01 = df_t01[df_t01['bron'] != 'https://id.erfgoed.net/themas/']
    df_t01 = df_t01[df_t01['bron'] != 'http://vocab.getty.edu/tgn/']
    df_t01 = df_t01[df_t01['bron'] != 'https://www.wikidata.org/entity/']
    df_t01 = df_t01[df_t01['bron'] != 'https://id.erfgoed.net/erfgoedobjecten/']
    return df_t01

# term komt meermaals voor
def ht02():
    df_t02 = df_thesaurus['term'].value_counts()
    df_t02 = df_t02.loc[lambda x: x > 1]
    df_t02 = pd.DataFrame({'term': df_t02.index, 'number of occurences': df_t02.values})
    return df_t02

# externe autoriteit komt meermaals voor
def ht03():
    df_t03 = df_thesaurus['term.nummer'].value_counts()
    df_t03 = df_t03.loc[lambda x: x > 1]
    df_t03 = pd.DataFrame({'term': df_t03.index, 'number of occurences': df_t03.values})
    return df_t03

# foutief wikidatanummer
def ht04():
    df_t04 = df_thesaurus[df_thesaurus['bron'] == 'https://www.wikidata.org/entity/']
    df_t04 = df_t04[~df_t04["term.nummer"].str.contains("Q", na=False)]
    return df_t04

#foutief AAT-nummer
def ht05():
    df_t05 = df_thesaurus[df_thesaurus['bron'] == 'http://vocab.getty.edu/aat/']
    df_t05 = df_t05[~df_t05['term.nummer'].isna()]
    df_t05 = df_t05['term.nummer'].astype(int)
    df_t05 = pd.DataFrame({'term': df_t05.values})
    df_t05 = df_t05[(df_t05['term'] > 999999999) | (df_t05['term'] < 100000000)]
    return df_t05

# foutief TGN-nummer
def ht06():
    df_t06 = df_thesaurus[df_thesaurus['bron'] == 'http://vocab.getty.edu/tgn/']
    df_t06 = df_t06[~df_t06['term.nummer'].isna()]
    df_t06 = df_t06['term.nummer'].astype(int)
    df_t06 = pd.DataFrame({'term': df_t06.values})
    df_t06 = df_t06[(df_t06['term'] > 9999999) | (df_t06['term'] < 1000000)]
    return df_t06

##################################################################################################################################################
# 10. R-SCHIJF
##################################################################################################################################################

# afbeeldingen op rschijf niet in adlib
def hr01():
    df_r01 = df_rschijflimited[~df_rschijflimited['objectnummer'].isin(df_collectie['objectnummer'])]
    return df_r01

# records in adlib niet op rschijf
def hr02():
    df_r02 = df_collectie[~df_collectie['objectnummer'].isin(df_rschijflimited['objectnummer'])]
    return df_r02

# ontbrekende afbeeldingen gevonden op rschijf
def hr03():
    df_a03 = df_collectie[df_collectie['reproductie.referentie'].isna()]
    df_r03 = df_a03[df_a03['objectnummer'].isin(df_rschijflimited['objectnummer'])]
    df_r03 = pd.merge(df_r03, df_rschijflimited, on="objectnummer", how='outer')
    df_r03 = df_r03[~df_r03['instelling.naam'].isna()]
    return df_r03

# _001 als deel van bestandsnaam
def hr04():
    df_r04 = df_rschijf[df_rschijf['objectnummer'].str.contains(r"_001", na=False)]
    return df_r04

# foutieve bestandsnamen
def hr05():
    df_r05 = df_rschijf[(df_rschijf['objectnummer'].str.contains("a", na=False))
                        | (df_rschijf['objectnummer'].str.contains("b", na=False))
                        | (df_rschijf['objectnummer'].str.contains("kopie", na=False))
                        | (df_rschijf['objectnummer'].str.contains(r"\(", na=False))
                        | (df_rschijf['objectnummer'].str.contains(r" 2", na=False))
                        | (df_rschijf['objectnummer'].str.contains(r"\)", na=False))
                        | (df_rschijf['objectnummer'].str.contains(r"c", na=False))
                        | (df_rschijf['objectnummer'].str.contains(r"d", na=False))
                        | (df_rschijf['objectnummer'].str.contains(r" 3", na=False))
                        | (df_rschijf['objectnummer'].str.contains(r" 4", na=False))
                        | (df_rschijf['objectnummer'].str.contains(r"C", na=False))]
    return df_r05

# foutieve start bestandsnamen
def hr06():
    df_r06 = df_rschijf[~((df_rschijf["objectnummer"].str.startswith("FO-", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("DB-", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("F0", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("AU-", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("19", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("20", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("DIA-", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("AF", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("RE-", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("VI-", na=False)))]
    return df_r06
