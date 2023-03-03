from django.shortcuts import render
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from django.http import HttpResponse
import datetime

#when deploying change to: df_collectie = pd.read_csv(r'/home/floreverkest/hva-im-dataportal/static/data/hva/collectie.csv', delimiter=';', low_memory=False)
df_collectie = pd.read_csv(r'dataportalapp\static\data\hva\collectie.csv', delimiter=';', low_memory=False)

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
    return response

def instellingsnaam(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="#001.xlsx"'
    wb = Workbook()
    if df_001.empty == True:
        ws = wb.active
        ws.title = '#001'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#002'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#003'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#004'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#005'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#006'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#007'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#008'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#009'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#010'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#011'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#012'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#013'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#014'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#015'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#016'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#017'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#018'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#019'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#020'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#021'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#022'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#023'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#024'
        ws['A1'] = "Empty Dataframe"
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
        ws = wb.active
        ws.title = '#025'
        ws['A1'] = "Empty Dataframe"
    else:
        ws = wb.active
        ws.title = '#025'
        rows = dataframe_to_rows(df_025, index=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(response)
    return response

def im(request):
    return render(request, 'im.html')
