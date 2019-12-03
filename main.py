from openpyxl import load_workbook
wb_val = load_workbook(filename='ExcelDocument.xlsx', data_only=True)
sheet_val = wb_val['Общая информация']
def Operation_capital():
    S_and_M_Per_Year = []
    NZP = []
    ZGP = []
    DZ = []
    KZ = []
    Diffe_value= []
    Summ=[]
    for i in range(5): #Заполнение массива с расчетами
        S_and_M_Per_Year.append((Max_production*1000*Obyom_proizv[i]*S_and_M)/Number_of_part[0])
        NZP.append(Max_production*1000*Obyom_proizv[i]*Sebe_stoim/Number_of_part[1])
        ZGP.append(Max_production*1000*Obyom_proizv[i]*Sebe_stoim/Number_of_part[2])
        DZ.append(Vyruchka_po_godam[i]/Number_of_part[3])
        KZ.append(Max_production*1000*Obyom_proizv[i]*S_and_M/Number_of_part[4])
    S_and_M_Per_Year.append(0)#Ввод последней строки с нулевыми значениями
    NZP.append(0)
    ZGP.append(0)
    DZ.append(0)
    KZ.append(0)
    for k in range(5):#Двумернй массив разности параметров
        Diffe_value.append([])
        for j in range(5):
            if j == 0:
                Diffe_value[k].append(S_and_M_Per_Year[k+1] - S_and_M_Per_Year[k])
            if j == 1:
                Diffe_value[k].append(NZP[k+1] - NZP[k])
            if j == 2:
                Diffe_value[k].append(ZGP[k+1] - ZGP[k])
            if j == 3:
                Diffe_value[k].append(DZ[k+1] - DZ[k])
            if j == 4:
                Diffe_value[k].append(KZ[k+1] - KZ[k])
    for k in range(5):
        Summ.append(Diffe_value[k][0]+Diffe_value[k][1]+Diffe_value[k][2]+Diffe_value[k][3]-Diffe_value[k][4])  

# Общая информация
Stoimost_oborud = sheet_val['B2'].value
Srok_pol_ispolz = sheet_val['B3'].value
Max_production = sheet_val['B4'].value
Akzion_finan_of_project = sheet_val['B5'].value
Trebue_doxodn = sheet_val['B6'].value
Stavka_po_kreditu = sheet_val['B7'].value
Stavka_po_nalogu = sheet_val['B8'].value
Nalogovaya_nagruzka = sheet_val['B9'].value
Amortization = 4000
index = ['B', 'C', 'D', 'E', 'F']
Obyom_proizv = []
Ostatochn_st_oborud = []
Pogashenie_dolg_kred=[]
for i in range(len(index)):
    Obyom_proizv.append(sheet_val[index[i]+'12'].value)
    Ostatochn_st_oborud.append(sheet_val[index[i]+'13'].value)
    Pogashenie_dolg_kred.append(sheet_val[index[i]+'14'].value)
#Доходы и расходы
Cost_per_unit = sheet_val['B18'].value
Sebe_stoim = sheet_val['B19'].value
S_and_M  = sheet_val['B21'].value
ZP = sheet_val['B22'].value
Others = sheet_val['B23'].value
Komm_rasx = sheet_val['B24'].value
Uprav_rasx = sheet_val['B25'].value
Summa_rasxodov = (Komm_rasx+Uprav_rasx)*1000



#Период оборота (в днях)
S_and_M_zapasy = sheet_val['B27'].value
Nezav_proizv=sheet_val['B28'].value
Zapasi_got_prod = sheet_val['B29'].value
Debet = sheet_val['B30'].value
Kredit = sheet_val['B31'].value
Day_per_year=sheet_val['B32'].value



#Подсчет количества оборотов в год
Number_of_part = []
for i in range (5):
    Number_of_part.append(Day_per_year/(sheet_val['B'+str(27+i)].value))

Vyruchka_po_godam = []
Vyruchka_po_godam_min_nalog=[]
Valovaya_pribyl = []
Pribyl_ot_prodazh = []
for i in range (5):
    Vyruchka_po_godam.append(Obyom_proizv[i]*Max_production*Cost_per_unit)
    Vyruchka_po_godam_min_nalog.append(Vyruchka_po_godam[i]*(1-Nalogovaya_nagruzka))
    Valovaya_pribyl.append(Vyruchka_po_godam_min_nalog[i]-(Max_production*Obyom_proizv[i]*(Cost_per_unit-Sebe_stoim))-Amortization)
    Pribyl_ot_prodazh.append(Valovaya_pribyl[i]-Summa_rasxodov)
Operation_capital()
