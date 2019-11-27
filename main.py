from openpyxl import load_workbook
wb_val = load_workbook(filename='ExcelDocument.xlsx', data_only=True)
sheet_val = wb_val['Общая информация']
Cost_per_unit = sheet_val['B18'].value
Stoimost_oborud = sheet_val['B2'].value
Srok_pol_ispolz = sheet_val['B3'].value
Max_production = sheet_val['B4'].value
Akzion_finan_of_project = sheet_val['B5'].value
Trebue_doxodn = sheet_val['B6'].value
Stavka_po_kreditu = sheet_val['B7'].value
Stavka_po_nalogu = sheet_val['B8'].value
Nalogovaya_nagruzka = sheet_val['B9'].value
index = ['B','C','D','E','F']
Obyom_proizv = []
Ostatochn_st_oborud = []
Pogashenie_dolg_kred=[]
for i in range(len(index)):
    Obyom_proizv.append(sheet_val[index[i]+'12'].value)
    Ostatochn_st_oborud.append(sheet_val[index[i]+'13'].value)
    Pogashenie_dolg_kred.append(sheet_val[index[i]+'14'].value)
print(Obyom_proizv)