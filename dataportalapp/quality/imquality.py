import pandas as pd
import datetime

df_collectie = pd.read_csv(r'C:\Users\Verkesfl\OneDrive - Groep Gent\Bureaublad\hva-im-dataportal\dataportalapp\static\data\im\collectie.csv', delimiter=';',
                           low_memory=False)
df_thesaurus = pd.read_csv(r'C:\Users\Verkesfl\OneDrive - Groep Gent\Bureaublad\hva-im-dataportal\dataportalapp\static\data\im\thesaurus.csv', delimiter=';', low_memory=False)

df_bron = pd.read_excel(r'C:\Users\Verkesfl\OneDrive - Groep Gent\Bureaublad\hva-im-dataportal\dataportalapp\static\data\im\bron.xlsx')
df_rschijf = pd.read_excel(r'C:\Users\Verkesfl\OneDrive - Groep Gent\Bureaublad\hva-im-dataportal\dataportalapp\static\data\im\rschijf.xlsx')
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains("WERKMAP", na=False)]
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains("EXTERNE SCHIJVEN", na=False)]
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains("FICHE", na=False)]
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains(r"A3", na=False)]
df_rschijf = df_rschijf[~df_rschijf['objectnummer'].str.contains(r"_", na=False)]
df_rschijflimited = df_rschijf[(df_rschijf["objectnummer"].str.startswith("F", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("D", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("V", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("AU", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("PK", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("AF", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("RE", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("VI", na=False))]
df_rschijflimited = df_rschijflimited[~df_rschijflimited['objectnummer'].str.contains(r"_", na=False)]
df_rschijflimited = df_rschijflimited[~((df_rschijflimited['objectnummer'].str.contains("a", na=False))
                    | (df_rschijf['objectnummer'].str.contains("b", na=False))
                    | (df_rschijf['objectnummer'].str.contains("kopie", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"\(", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r" 2", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"\)", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"c", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"d", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"e", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"B", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"f", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r" 3", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r" 4", na=False))
                    | (df_rschijf['objectnummer'].str.contains(r"C", na=False)))]

year = datetime.datetime.now().year

df_collectie['wijziging.naam'] = df_collectie['wijziging.naam'].str.split('$').str[0]
df_collectie['wijziging.datum'] = df_collectie['wijziging.datum'].str.split('$').str[0]

##################################################################################################################################################
# 1. COLLECTIE
##################################################################################################################################################

# foutieve instellingsnaam
def i001():
    df_001 = df_collectie[df_collectie["instelling.naam"] != 'Industriemuseum']
    return df_001

# geen collectie
def i002():
    df_002 = df_collectie[df_collectie['collectie'].isna()]
    return df_002

# objectnummer foutieve start
def i003():
    df_003 = df_collectie[~df_collectie['objectnummer'].str.startswith(('AU', 'F', 'D', 'DC', 'V', 'VI', 'AF', 'RE', 'PK'))]
    return df_003

# objectnummer foutieve format:
def i004():
    df_01 = df_collectie[~df_collectie['objectnummer'].apply(lambda x: len(str(x)) == 7)]
    df_02 = df_01[~df_01['objectnummer'].apply(lambda x: len(str(x)) == 6)]
    df_03 = df_02[~df_02['objectnummer'].apply(lambda x: len(str(x)) == 10)]
    df_004 = df_03[~df_03['objectnummer'].apply(lambda x: len(str(x)) == 11)]
    return df_004

# lege basisregistratie
def i005():
    df_005 = df_collectie[df_collectie['objectnaam'].isna()]
    return df_005

def i006():
    df_006 = df_collectie[df_collectie['titel'].isna()]
    return df_006

def i007():
    df_007 = df_collectie[df_collectie['reproductie.referentie'].isna()]
    return df_007

def i008():
    df_008 = df_collectie[df_collectie['associatie.onderwerp'].isna()]
    return df_008

def i009():
    df_009 = df_collectie[df_collectie['associatie.onderwerp.soort'].isna()]
    return df_009

# vervaardiging aanwezig maar geen asso periode
def i010():
    df_010 = df_collectie[df_collectie['vervaardiging.datum.begin'].notna()]
    df_010 = df_010[df_010['associatie.periode'].isna()]
    return df_010

# datum vervaardiging eind > datum vervaardiging begin
def i011():
    df_011 = df_collectie.loc[df_collectie['vervaardiging.datum.begin'] > df_collectie['vervaardiging.datum.eind']]
    return df_011

# foutief datum formaat
def i012():
    df_012 = df_collectie[df_collectie['vervaardiging.datum.begin'].notna()]
    df_012 = df_012.drop(df_012[pd.to_datetime(df_012['vervaardiging.datum.begin'], format='%Y-%m',
                                            errors='coerce').notna()].index)
    df_012 = df_012[~df_012['vervaardiging.datum.begin'].str.startswith('14')]
    df_012 = df_012[~df_012['vervaardiging.datum.begin'].str.startswith('15')]
    df_012 = df_012[~df_012['vervaardiging.datum.begin'].str.startswith('16')]
    return df_012

# precisie datum vervaardiging begin foutief
def i013():
    df_013 = df_collectie[df_collectie['vervaardiging.datum.begin.prec'].notna()]
    df_013 = df_013[df_013['vervaardiging.datum.begin.prec'] != 'ca.']
    df_013 = df_013[df_013['vervaardiging.datum.begin.prec'] != '$']
    df_013 = df_013[df_013['vervaardiging.datum.begin.prec'] != 'vanaf']
    df_013 = df_013[df_013['vervaardiging.datum.begin.prec'] != 'na']
    df_013 = df_013[df_013['vervaardiging.datum.begin.prec'] != 'voor']
    return df_013

# precisie datum vervaardiging eind foutief
def i014():
    df_014 = df_collectie[df_collectie['vervaardiging.datum.eind.prec'].notna()]
    df_014 = df_014[df_014['vervaardiging.datum.eind.prec'] != '$']
    df_014 = df_014[df_014['vervaardiging.datum.eind.prec'] != 'ca.']
    df_014 = df_014[df_014['vervaardiging.datum.eind.prec'] != 'voor']
    return df_014

# ontbrekende afmetingen
def i015():
    df_015 = df_collectie[df_collectie['afmeting.waarde'].isna()]
    return df_015

# ontbrekend materiaal
def i016():
    df_016 = df_collectie[df_collectie['materiaal'].isna()]
    return df_016

# ontbrekende rechtenstatus
def i017():
    df_017 = df_collectie[df_collectie['rechten.type'].isna()]
    return df_017

# ontbrekende machtigingsstatus
def i018():
    df_018 = df_collectie[~df_collectie['rechten.type'].isna()]
    df_018 = df_018[~df_018['rechten.machtigingsstatus'].str.contains("toegewezen", na=False)]
    df_018 = df_018[~df_018['rechten.machtigingsstatus'].str.contains("berekend risico", na=False)]
    return df_018

# ontbrekend formulier
def i019():
    df_019 = df_collectie[(df_collectie['rechten.type'] == "IN COPYRIGHT - NON-COMMERCIAL USE PERMITTED") |
                        (df_collectie['rechten.type'] == "CC-BY-NC 4.0") |
                        (df_collectie['rechten.type'] == "CC-BY-SA 4.0")]
    df_019 = df_019[df_019['rechten.referentienummer'].isna()]
    return df_019

# records in publiek domein
def i020():
    df_020 = df_collectie[df_collectie['rechten.type'] != 'PUBLIC DOMAIN']
    df_020_1 = df_020[df_020['vervaardiging.datum.eind'].isna()]
    df_020_1 = df_020_1[~df_020_1['vervaardiging.datum.begin'].isna()]
    df_020_1['vervaardiging.datum.begin'] = df_020_1['vervaardiging.datum.begin'].astype(str)
    df_020_1['vervaardiging.datum.begin'] = df_020_1['vervaardiging.datum.begin'].str[:4]
    df_020_1['vervaardiging.datum.begin'] = df_020_1['vervaardiging.datum.begin'].astype(int)
    df_020_1 = df_020_1[df_020_1['vervaardiging.datum.begin'] <= (year - 150)]

    df_020_2 = df_020[~df_020['vervaardiging.datum.eind'].isna()]
    df_020_2['vervaardiging.datum.eind'] = df_020_2['vervaardiging.datum.eind'].astype(str)
    df_020_2['vervaardiging.datum.eind'] = df_020_2['vervaardiging.datum.eind'].str[:4]
    df_020_2 = df_020_2[df_020_2['vervaardiging.datum.eind'] != '$']
    df_020_2['vervaardiging.datum.eind'] = df_020_2['vervaardiging.datum.eind'].astype(int)
    df_020_2 = df_020_2[df_020_2['vervaardiging.datum.eind'] <= (year - 150)]

    frames = [df_020_1, df_020_2]
    df_020 = pd.concat(frames)
    return df_020

# ontbrekende uitleg waarom publiek domein bij bijzonderheden
def i021():
    df_021 = df_collectie[df_collectie['rechten.type'] == 'PUBLIC DOMAIN']
    df_021 = df_021[df_021['rechten.bijzonderheden'].isna()]
    return df_021

# TO DO!!! filteren op objectnaam type: born digital collecties
#foutieve toestand
def i022():
    df_022 = df_collectie[df_collectie['objectnaam.type'] != 'born digital collecties']
    df_022 = df_022[~df_022['toestand'].str.contains("goed", na=False)]
    df_022 = df_022[~df_022['toestand'].str.contains("matig", na=False)]
    df_022 = df_022[~df_022['toestand'].str.contains("slecht", na=False)]
    df_022 = df_022[~df_022['toestand'].str.contains("redelijk", na=False)]
    return df_022

# ontbrekende/foutieve verwervingsmethode
def i023():
    df_023 = df_collectie[df_collectie['verwerving.methode'] != 'schenking']
    df_023 = df_023[df_023['verwerving.methode'] != 'aankoop']
    df_023 = df_023[df_023['verwerving.methode'] != 'onbekend']
    df_023 = df_023[df_023['verwerving.methode'] != 'bruikleen']
    df_023 = df_023[df_023['verwerving.methode'] != 'legaat']
    df_023 = df_023[df_023['verwerving.methode'] != 'overdracht']
    df_023 = df_023[df_023['verwerving.methode'] != 'eigen productie']
    df_023 = df_023[df_023['verwerving.methode'] != 'vondst']
    return df_023

# titel eindigt op .
def i024():
    df_024 = df_collectie[df_collectie['titel'].str.endswith(".", na=False)]
    return df_024

# titel bevat gent, maar geen associatie gent
def i025():
    df_025 = df_collectie[df_collectie["titel"].str.contains("Gent", na=False)]
    df_025 = df_025[~df_025["associatie.onderwerp"].str.contains("Gent", na=False)]
    return df_025

# titel start niet met hoofdletter
def i026():
    df_026 = df_collectie[~df_collectie['titel'].isna()]
    df_026['t'] = df_026['titel'].astype(str).str[0]
    df_026 = df_026[~df_026['t'].str.isupper()]
    df_026 = df_026[~df_026['t'].str.isdigit()]
    df_026 = df_026[~df_026['titel'].str.startswith("'s ", na=False)]
    df_026 = df_026[~df_026['titel'].str.startswith("'t ", na=False)]
    return df_026

# titel heeft ( of ) staan
def i027():
    search = [r"\)", r"\("]
    df_027 = df_collectie[df_collectie['titel'].str.contains('|'.join(search), na=False)]
    return df_027

##################################################################################################################################################
# 2. THESAURUS
##################################################################################################################################################

# bron afwezig of niet correct
def it01():
    df_t01 = df_thesaurus[df_thesaurus['bron'] != 'http://vocab.getty.edu/aat/']
    df_t01 = df_t01[df_t01['bron'] != 'https://id.erfgoed.net/themas/']
    df_t01 = df_t01[df_t01['bron'] != 'http://vocab.getty.edu/tgn/']
    df_t01 = df_t01[df_t01['bron'] != 'https://www.wikidata.org/entity/']
    df_t01 = df_t01[df_t01['bron'] != 'https://id.erfgoed.net/erfgoedobjecten/']
    return df_t01

# term komt meermaals voor
def it02():
    df_t02 = df_thesaurus['term'].value_counts()
    df_t02 = df_t02.loc[lambda x: x > 1]
    df_t02 = pd.DataFrame({'term': df_t02.index, 'number of occurences': df_t02.values})
    return df_t02

# externe autoriteit komt meermaals voor
def it03():
    df_t03 = df_thesaurus['term.nummer'].value_counts()
    df_t03 = df_t03.loc[lambda x: x > 1]
    df_t03 = pd.DataFrame({'term': df_t03.index, 'number of occurences': df_t03.values})
    return df_t03

# foutief wikidatanummer
def it04():
    df_t04 = df_thesaurus[df_thesaurus['bron'] == 'https://www.wikidata.org/entity/']
    df_t04 = df_t04[~df_t04["term.nummer"].str.contains("Q", na=False)]
    return df_t04

#foutief AAT-nummer
def it05():
    df_t05 = df_thesaurus[df_thesaurus['bron'] == 'http://vocab.getty.edu/aat/']
    df_t05 = df_t05[~df_t05['term.nummer'].isna()]
    df_t05 = df_t05['term.nummer'].astype(int)
    df_t05 = pd.DataFrame({'term': df_t05.values})
    df_t05 = df_t05[(df_t05['term'] > 999999999) | (df_t05['term'] < 100000000)]
    return df_t05

# foutief TGN-nummer
def it06():
    df_t06 = df_thesaurus[df_thesaurus['bron'] == 'http://vocab.getty.edu/tgn/']
    df_t06 = df_t06[~df_t06['term.nummer'].isna()]
    df_t06 = df_t06['term.nummer'].astype(int)
    df_t06 = pd.DataFrame({'term': df_t06.values})
    df_t06 = df_t06[(df_t06['term'] > 9999999) | (df_t06['term'] < 1000000)]
    return df_t06

##################################################################################################################################################
# 10. R-SCHIJF
##################################################################################################################################################

#afbeeldingen op rschijf niet in adlib
def ir01():
    df_r01 = df_rschijflimited[~df_rschijflimited['objectnummer'].isin(df_collectie['objectnummer'])]
    return df_r01

#afbeeldingen in adlib niet op rschijf
def ir02():
    df_r02 = df_collectie[~df_collectie['objectnummer'].isin(df_rschijflimited['objectnummer'])]
    return df_r02

#ontbrekende afbeeldingen adlib, gevonden op rschijf
def ir03():
    df_a03 = df_collectie[df_collectie['reproductie.referentie'].isna()]
    df_r03 = df_a03[df_a03['objectnummer'].isin(df_rschijflimited['objectnummer'])]
    df_r03 = pd.merge(df_r03, df_rschijf, on="objectnummer", how='outer')
    df_r03 = df_r03[~df_r03['instelling.naam'].isna()]
    return df_r03

#foutieve bestandsnaam
def ir04():
    df_r04 = df_rschijf[(df_rschijf['objectnummer'].str.contains("a", na=False))
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
    return df_r04

# foutieve start bestandsnamen
def ir05():
    df_r05 = df_rschijf[~((df_rschijf["objectnummer"].str.startswith("F", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("D", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("AU", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("AF", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("PK", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("RE", na=False))
                            | (df_rschijf["objectnummer"].str.startswith("V", na=False)))]
    return df_r05

# bestand in map bron, niet in map collectie
def ir06():
    df_r06 = df_bron[~df_bron['objectnummer'].isin(df_rschijf['objectnummer'])]
    return df_r06

