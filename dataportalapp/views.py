from django.shortcuts import render
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from django.http import HttpResponse
import datetime
from django.contrib import messages


#when deploying change to: df_collectie = pd.read_csv(r'/home/floreverkest/hva-im-dataportal/static/data/hva/collectie.csv', delimiter=';', low_memory=False)
df_collectie = pd.read_csv(r'dataportalapp\static\data\hva\collectie.csv', delimiter=';', low_memory=False)
df_thesaurus = pd.read_csv(r'C:\Users\Verkesfl\OneDrive - Groep Gent\Bureaublad\hva-im-dataportal\dataportalapp\static\data\hva\thesaurus.csv', delimiter=';', low_memory=False)
df_rschijf = pd.read_excel(r'C:\Users\Verkesfl\OneDrive - Groep Gent\Bureaublad\hva-im-dataportal\dataportalapp\static\data\hva\rschijf.xlsx')

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

response = HttpResponse(content_type='application/ms-excel')
response['Content-Disposition'] = 'attachment; filename="#000.xlsx"'

df_001 = df_collectie[df_collectie["instelling.naam"] != 'Het Huis van Alijn (Gent)']

df_002 = df_collectie[df_collectie["instelling.code"] != 'INST-570']

# objectnummer foutieve start
df_003 = df_collectie[~df_collectie['objectnummer'].str.startswith(('AU-', 'FO-', '19', '20', 'DIA-', 'AF', 'DB-', 'RE-',
                                                            'F0', 'VI'))]

# objectnummer foutieve format:

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

df_005 = df_collectie[df_collectie['onderscheidende_kenmerken'] != 'DIGITALE COLLECTIE']
df_005 = df_005[df_005['onderscheidende_kenmerken'] != 'OBJECT']
df_005 = df_005[df_005['onderscheidende_kenmerken'] != 'BEELD']
df_005 = df_005[df_005['onderscheidende_kenmerken'] != 'DOCUMENTAIRE COLLECTIE']

df_006 = df_collectie[df_collectie['objectnaam'].isna()]
df_007 = df_collectie[df_collectie['titel'].isna()]
df_008 = df_collectie[df_collectie['reproductie.referentie'].isna()]
df_009 = df_collectie[df_collectie['associatie.onderwerp'].isna()]

df_010 = df_collectie[df_collectie['vervaardiging.plaats'].notna()]
df_010 = df_010[df_010['vervaardiging.plaats'] != '$']
df_010['vervaardiging.plaats'] = df_010['vervaardiging.plaats'].map(lambda x: x.lstrip('$').rstrip('$'))
df_010 = df_010[~df_010["vervaardiging.plaats"].isin(df_010["associatie.onderwerp"])]

df_011 = df_collectie[df_collectie['vervaardiging.datum.begin'].notna()]
df_011 = df_011[df_011['associatie.periode'].isna()]

df_012 = df_collectie.loc[df_collectie['vervaardiging.datum.begin'] == df_collectie['vervaardiging.datum.eind']]

df_013 = df_collectie.loc[df_collectie['vervaardiging.datum.begin'] > df_collectie['vervaardiging.datum.eind']]

df_014 = df_collectie[df_collectie['vervaardiging.datum.begin'].notna()]
df_014 = df_014.drop(df_014[pd.to_datetime(df_014['vervaardiging.datum.begin'], format='%Y-%m',
                                        errors='coerce').notna()].index)
df_014 = df_014[~df_014['vervaardiging.datum.begin'].str.startswith('14')]
df_014 = df_014[~df_014['vervaardiging.datum.begin'].str.startswith('15')]
df_014 = df_014[~df_014['vervaardiging.datum.begin'].str.startswith('16')]

df_015 = df_collectie[df_collectie['afmeting.waarde'].isna()]

df_016 = df_collectie[df_collectie['onderscheidende_kenmerken'] == 'OBJECT']
df_016 = df_016[~df_016['afmeting.eenheid'].isna()]
df_016 = df_016[~df_016['afmeting.eenheid'].str.startswith('cm')]

df_017 = df_collectie[df_collectie['onderscheidende_kenmerken'] == 'DIGITALE COLLECTIE']
df_017 = df_017[~df_017['afmeting.eenheid'].isna()]
df_017 = df_017[~df_017['afmeting.eenheid'].str.startswith('min')]
df_017 = df_017[~df_017['afmeting.eenheid'].str.startswith('kB')]
df_017 = df_017[~df_017['afmeting.eenheid'].str.startswith('MB')]
df_017 = df_017[~df_017['afmeting.eenheid'].str.startswith('GB')]

df_018 = df_collectie[df_collectie['onderscheidende_kenmerken'] == 'DOCUMENTAIRE COLLECTIE|BEELD']
df_018 = df_018[~df_018['afmeting.eenheid'].isna()]
df_018 = df_018[~df_018['afmeting.eenheid'].str.startswith('mm')]

df_019 = df_collectie[df_collectie['rechten.type'].isna()]

df_020 = df_collectie[~df_collectie['rechten.type'].isna()]
df_020 = df_020[~df_020['rechten.machtigingsstatus'].str.contains("toegewezen", na=False)]

df_021 = df_collectie[(df_collectie['rechten.type'] == "IN COPYRIGHT - NON-COMMERCIAL USE PERMITTED") |
                (df_collectie['rechten.type'] == "CC-BY-NC 4.0") |
                (df_collectie['rechten.type'] == "CC-BY-SA 4.0")]
df_021 = df_021[df_021['rechten.referentienummer'].isna()]

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

df_023 = df_collectie[~df_collectie['toestand'].str.contains("goed", na=False)]
df_023 = df_023[~df_023['toestand'].str.contains("matig", na=False)]
df_023 = df_023[~df_023['toestand'].str.contains("slecht", na=False)]

df_024 = df_collectie[df_collectie['verwerving.methode'] != 'schenking']
df_024 = df_024[df_024['verwerving.methode'] != 'aankoop']
df_024 = df_024[df_024['verwerving.methode'] != 'onbekend']
df_024 = df_024[df_024['verwerving.methode'] != 'bruikleen']

df_025 = df_collectie[df_collectie["titel"].str.contains("Gent", na=False)]
df_025 = df_025[~df_025["associatie.onderwerp"].str.contains("Gent", na=False)]

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

# 9. THESAURUS

# bron afwezig of niet correct
df_t01 = df_thesaurus[df_thesaurus['bron'] != 'http://vocab.getty.edu/aat/']
df_t01 = df_t01[df_t01['bron'] != 'https://id.erfgoed.net/themas/']
df_t01 = df_t01[df_t01['bron'] != 'http://vocab.getty.edu/tgn/']
df_t01 = df_t01[df_t01['bron'] != 'https://www.wikidata.org/entity/']
df_t01 = df_t01[df_t01['bron'] != 'https://id.erfgoed.net/erfgoedobjecten/']

# term komt meermaals voor
df_t02 = df_thesaurus['term'].value_counts()
df_t02 = df_t02.loc[lambda x: x > 1]
df_t02 = pd.DataFrame({'term': df_t02.index, 'number of occurences': df_t02.values})

# externe autoriteit komt meermaals voor
df_t03 = df_thesaurus['term.nummer'].value_counts()
df_t03 = df_t03.loc[lambda x: x > 1]
df_t03 = pd.DataFrame({'term': df_t03.index, 'number of occurences': df_t03.values})

# foutief wikidatanummer
df_t04 = df_thesaurus[df_thesaurus['bron'] == 'https://www.wikidata.org/entity/']
df_t04 = df_t04[~df_t04["term.nummer"].str.contains("Q", na=False)]

#foutief AAT-nummer
df_t05 = df_thesaurus[df_thesaurus['bron'] == 'http://vocab.getty.edu/aat/']
df_t05 = df_t05[~df_t05['term.nummer'].isna()]
df_t05 = df_t05['term.nummer'].astype(int)
df_t05 = pd.DataFrame({'term': df_t05.values})
df_t05 = df_t05[(df_t05['term'] > 999999999) | (df_t05['term'] < 100000000)]

# foutief TGN-nummer
df_t06 = df_thesaurus[df_thesaurus['bron'] == 'http://vocab.getty.edu/tgn/']
df_t06 = df_t06[~df_t06['term.nummer'].isna()]
df_t06 = df_t06['term.nummer'].astype(int)
df_t06 = pd.DataFrame({'term': df_t06.values})
df_t06 = df_t06[(df_t06['term'] > 9999999) | (df_t06['term'] < 1000000)]

df_rschijf = df_rschijf[(df_rschijf["objectnummer"].str.startswith("FO-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("DB-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("F0", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("AU-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("19", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("20", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("DIA-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("AF", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("RE-", na=False))
                        | (df_rschijf["objectnummer"].str.startswith("VI-", na=False))]

df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains("WERKMAP", na=False)]
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains("EXTERNE SCHIJVEN", na=False)]
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains(r"_", na=False)]
df_rschijf = df_rschijf[~df_rschijf['pad'].str.contains(r"A3", na=False)]
df_r04 = df_rschijf[df_rschijf['objectnummer'].str.contains(r"_001", na=False)]
df_rschijf = df_rschijf[~df_rschijf['objectnummer'].str.contains(r"_", na=False)]
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
df_rschijf = df_rschijf[~((df_rschijf['objectnummer'].str.contains("a", na=False))
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

df_r01 = df_rschijf[~df_rschijf['objectnummer'].isin(df_collectie['objectnummer'])]
df_r02 = df_collectie[~df_collectie['objectnummer'].isin(df_rschijf['objectnummer'])]
df_a03 = df_collectie[df_collectie['reproductie.referentie'].isna()]
df_r03 = df_a03[df_a03['objectnummer'].isin(df_rschijf['objectnummer'])]
df_r03 = pd.merge(df_r03, df_rschijf, on="objectnummer", how='outer')
df_r03 = df_r03[~df_r03['instelling.naam'].isna()]

# Create your views here.
def home(request):
    return render(request, 'home.html')

def hva(request):
    return render(request, 'hva.html')

def all(request):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Info'
    ws['A1'] = "list of sheet tab codes"
    ws.append(['sheet number', 'quality check'])
    ws.append(['#001', 'instellingsnaam != Het Huis van Alijn (Gent)'])
    ws.append(['#002', 'instellingscode != INST-570'])
    ws.append(['#003', 'objectnummer start niet met: AU-, FO-, 19, 20, DIA-, AF, DB-, RE-, F0, VI-'])
    ws.append(['#004', 'lengte objectnummer niet correct'])
    ws.append(['#005', 'foutief of leeg onderscheidende kenmerken'])
    ws.append(['#006', 'ontbrekende objectnaam'])
    ws.append(['#007', 'ontbrekende titel'])
    ws.append(['#008', 'ontbrekende afbeelding'])
    ws.append(['#009', 'ontbrekende associatie onderwerp'])
    ws.append(['#010', 'vervaardiging plaats niet in associatie onderwerp'])
    ws.append(['#011', 'vervaardiging datering niet in associatie periode'])
    ws.append(['#012', 'vervaardiging datum begin == vervaardiging datum eind'])
    ws.append(['#013', 'vervaardiging datum eind > datum begin'])
    ws.append(['#014', 'geen datering format'])
    ws.append(['#015', 'ontbrekende afmeting'])
    ws.append(['#016', 'objecten afmeting niet in cm'])
    ws.append(['#017', 'digitale collectie afmetingen niet in min, kb, mb of gb'])
    ws.append(['#018', 'documentaire collectie, afmetingen niet in mm'])
    ws.append(['#019', 'ontbrekende rechtenstatus'])
    ws.append(['#020', 'ontbrekende machtigingsstatus'])
    ws.append(['#021', 'ontbrekend formulier'])
    ws.append(['#022', 'publiek domein'])
    ws.append(['#023', 'ontbrekende/foutieve toestand'])
    ws.append(['#024', 'ontbrekende/foutieve verwervingsmethode'])
    ws.append(['#025', 'titel bevat Gent AND NOT associatie.onderwerp = Gent'])
    ws.append(['#026', 'associatie wereldtentoonstelling Brussel, Antwerpen, Gent, Luik, Parijs, Amsterdam, Londen en niet stad als associatie'])
    ws.append(['#t01', 'thesaurus bron afwezig of niet correct'])
    ws.append(['#t02', 'thesaurus term komt meermaals voor'])
    ws.append(['#t03', 'thesaurus externe autoriteit komt meermaals voor'])
    ws.append(['#t04', 'thesaurus foutief wikidatanummer'])
    ws.append(['#t05', 'thesaurus foutief AAT-nummer'])
    ws.append(['#t06', 'thesaurus foutief TGN-nummer'])

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 60

    if df_001.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#001")
        rows = dataframe_to_rows(df_001, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_002.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#002")
        rows = dataframe_to_rows(df_002, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_003.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#003")
        rows = dataframe_to_rows(df_003, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_004.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#004")
        rows = dataframe_to_rows(df_004, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_005.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#005")
        rows = dataframe_to_rows(df_005, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_006.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#006")
        rows = dataframe_to_rows(df_006, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_007.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#007")
        rows = dataframe_to_rows(df_007, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_008.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#008")
        rows = dataframe_to_rows(df_008, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_009.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#009")
        rows = dataframe_to_rows(df_009, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_010.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#010")
        rows = dataframe_to_rows(df_010, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_011.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#011")
        rows = dataframe_to_rows(df_011, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_012.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#012")
        rows = dataframe_to_rows(df_012, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_013.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#013")
        rows = dataframe_to_rows(df_013, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_014.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#014")
        rows = dataframe_to_rows(df_014, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_015.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#015")
        rows = dataframe_to_rows(df_015, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_016.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#016")
        rows = dataframe_to_rows(df_016, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_017.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#017")
        rows = dataframe_to_rows(df_017, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_018.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#018")
        rows = dataframe_to_rows(df_018, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_019.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#019")
        rows = dataframe_to_rows(df_019, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_020.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#020")
        rows = dataframe_to_rows(df_020, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_021.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#021")
        rows = dataframe_to_rows(df_021, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_022.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#022")
        rows = dataframe_to_rows(df_022, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_023.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#023")
        rows = dataframe_to_rows(df_023, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_024.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#024")
        rows = dataframe_to_rows(df_024, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_025.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#025")
        rows = dataframe_to_rows(df_025, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    if df_026.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#026")
        rows = dataframe_to_rows(df_026, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_t01.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#T01")
        rows = dataframe_to_rows(df_t01, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_t02.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#T02")
        rows = dataframe_to_rows(df_t02, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_t03.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#T03")
        rows = dataframe_to_rows(df_t03, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_t04.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#T04")
        rows = dataframe_to_rows(df_t04, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_t05.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#T05")
        rows = dataframe_to_rows(df_t05, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    if df_t06.empty == True:
        print('empty dataframe')
    else:
        ws = wb.create_sheet("#T06")
        rows = dataframe_to_rows(df_t06, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def instellingsnaam(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#001.xlsx"'
    wb = Workbook()
    if df_001.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#001'
        rows = dataframe_to_rows(df_001, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def instellingscode(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#002.xlsx"'
    wb = Workbook()
    if df_002.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#002'
        rows = dataframe_to_rows(df_002, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def objectnummer(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#003.xlsx"'
    wb = Workbook()
    if df_003.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#003'
        rows = dataframe_to_rows(df_003, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def objectnmr(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#004.xlsx"'
    wb = Workbook()
    if df_004.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#004'
        rows = dataframe_to_rows(df_004, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def onderscheidendkenmerk(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#005.xlsx"'
    wb = Workbook()
    if df_005.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#005'
        rows = dataframe_to_rows(df_005, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def objectnaam(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#006.xlsx"'
    wb = Workbook()
    if df_006.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#006'
        rows = dataframe_to_rows(df_006, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def titel(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#007.xlsx"'
    wb = Workbook()
    if df_007.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#007'
        rows = dataframe_to_rows(df_007, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def afbeelding(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#008.xlsx"'
    wb = Workbook()
    if df_008.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#008'
        rows = dataframe_to_rows(df_008, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def associatie(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#009.xlsx"'
    wb = Workbook()
    if df_009.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#009'
        rows = dataframe_to_rows(df_009, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def associatieplaats(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#010.xlsx"'
    wb = Workbook()
    if df_010.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#010'
        rows = dataframe_to_rows(df_010, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def associatieperiode(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#011.xlsx"'
    wb = Workbook()
    if df_011.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#011'
        rows = dataframe_to_rows(df_011, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def datum(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#012.xlsx"'
    wb = Workbook()
    if df_012.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#012'
        rows = dataframe_to_rows(df_012, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def datumgroter(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#013.xlsx"'
    wb = Workbook()
    if df_013.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#013'
        rows = dataframe_to_rows(df_013, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def datumformat(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#014.xlsx"'
    wb = Workbook()
    if df_014.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#014'
        rows = dataframe_to_rows(df_014, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def afmeting(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#015.xlsx"'
    wb = Workbook()
    if df_015.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#015'
        rows = dataframe_to_rows(df_015, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def afmetingo(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#016.xlsx"'
    wb = Workbook()
    if df_016.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#016'
        rows = dataframe_to_rows(df_016, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def afmetingd(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#017.xlsx"'
    wb = Workbook()
    if df_017.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#017'
        rows = dataframe_to_rows(df_017, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def afmetingdd(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#018.xlsx"'
    wb = Workbook()
    if df_018.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#018'
        rows = dataframe_to_rows(df_018, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
        wb.save(response)
        return response

def rechten(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#019.xlsx"'
    wb = Workbook()
    if df_019.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#019'
        rows = dataframe_to_rows(df_019, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def rechtentype(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#020.xlsx"'
    wb = Workbook()
    if df_020.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#020'
        rows = dataframe_to_rows(df_020, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def rechtenref(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#021.xlsx"'
    wb = Workbook()
    if df_021.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#021'
        rows = dataframe_to_rows(df_021, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def pd(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#022.xlsx"'
    wb = Workbook()
    if df_022.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#022'
        rows = dataframe_to_rows(df_022, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def toestand(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#023.xlsx"'
    wb = Workbook()
    if df_023.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#023'
        rows = dataframe_to_rows(df_023, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def verwerving(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#024.xlsx"'
    wb = Workbook()
    if df_024.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#024'
        rows = dataframe_to_rows(df_024, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def variatitel(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#025.xlsx"'
    wb = Workbook()
    if df_025.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#025'
        rows = dataframe_to_rows(df_025, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def wereldtentoonstelling(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#026.xlsx"'
    wb = Workbook()
    if df_026.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#026'
        rows = dataframe_to_rows(df_026, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def bron(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#T01.xlsx"'
    wb = Workbook()
    if df_t01.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#T01'
        rows = dataframe_to_rows(df_t01, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def term(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#T02.xlsx"'
    wb = Workbook()
    if df_t02.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#T02'
        rows = dataframe_to_rows(df_t02, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def externeautoriteit(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#T03.xlsx"'
    wb = Workbook()
    if df_t03.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#T03'
        rows = dataframe_to_rows(df_t03, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def wikidata(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#T04.xlsx"'
    wb = Workbook()
    if df_t04.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#T04'
        rows = dataframe_to_rows(df_t04, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def aat(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#T05.xlsx"'
    wb = Workbook()
    if df_t05.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#T05'
        rows = dataframe_to_rows(df_t05, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def tgn(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#T06.xlsx"'
    wb = Workbook()
    if df_t06.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#T06'
        rows = dataframe_to_rows(df_t06, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def rschijf(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#R01.xlsx"'
    wb = Workbook()
    if df_r01.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#R01'
        rows = dataframe_to_rows(df_r01, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def adlib(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#R02.xlsx"'
    wb = Workbook()
    if df_r02.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#R02'
        rows = dataframe_to_rows(df_r02, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def bestandsnaam(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#R05.xlsx"'
    wb = Workbook()
    if df_r05.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#R05'
        rows = dataframe_to_rows(df_r05, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def start(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#R04.xlsx"'
    wb = Workbook()
    if df_r04.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#R04'
        rows = dataframe_to_rows(df_r04, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def afad(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#R03.xlsx"'
    wb = Workbook()
    if df_r03.empty == True:
        messages.success(request, 'Congrats! Empty list :)')
        return render(request, 'hva.html')
    else:
        ws = wb.active
        ws.title = '#R03'
        rows = dataframe_to_rows(df_r03, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def im(request):
    return render(request, 'im.html')

