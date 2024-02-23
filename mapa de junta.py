import xlsxwriter
limitador=input("digite start para começar")
l=[]
weld=[]
welder=[]
mat1=[]
mat2=[]
w=[]
wdate=[]
ndate=[]
v=[]
lp=[]
us=[]
project=input("digite o número do projeto")
project=str(project)
client=input("digite o cliente")
num_project=input("digite o número do projeto")
num_project=str(num_project)
drawing=input("digite o número do desenho")
drawing=str(drawing)
num_document=input("digite o número do documento")
num_document=str(num_document)
revision=input("digite a revisão")
revision=str(revision)
date=input("digite o data no seguinte formato xx/yy/zz")
date=str(date)
prepared_by=input("digite seu nome (emissor)")
note=input("digite a sua observação")
note=str(note)
def dados():
    weld_id=input("digite o ID da solda")
    weld_id=str(weld_id)
    weld.append(weld_id)
    material_1=input("digite o material 1")
    material_1=str(material_1)
    mat1.append(material_1)
    material_2=input("digite o material 2")
    material_2=str(material_2)
    mat2.append(material_2)
    wps=input("digite o WPS")
    wps=str(wps)
    w.append(wps)
    welder_id=input("digite o ID do soldador")
    welder_id=str(welder_id)
    welder.append(welder_id)
    weld_date=input("digite a data da soldagem no seguinte formato xx/yy/zz")
    weld_date=str(weld_date)
    wdate.append(weld_date)
    nde_date=input("digite a data do NDE no seguinte formato xx/yy/zz")
    nde_date=str(nde_date)
    ndate.append(nde_date)
    vs=input("digite o número do VS")
    vs=str(vs)
    v.append(vs)
    lp_pm=input("digite o número do LP/PM")
    lp_pm=str(lp_pm)
    lp.append(lp_pm)
    us_rx=input("digite o número do US/RX")
    us_rx=str(us_rx)
    us.append(us_rx)
while limitador=="start":
    breaker=input("digite sim para adicionar mais uma informação e parar para parar de adicionar")
    breaker=breaker.upper()
    breaker=breaker.lower()
    if breaker=="sim":
        dados()
    if breaker!="parar" and breaker!="sim":
        print("comando inválido, retorne")
        dados()
    if breaker=="parar":
        break
workbook=xlsxwriter.Workbook("Mapa de Junta.xlsx")
worksheet=workbook.add_worksheet("Mapa de junta RBA")
border_simple=workbook.add_format({
    "border": 2,
})
title_format = workbook.add_format({
    "bold": True,
    "border":   2,
    "align":    "center",
    "valign":   "vcenter",
    "fg_color":    "#00FFFF"    
})
border_format=workbook.add_format({
    "bold": True,    
    "border":   2,
    "align":    "center",
    "valign":   "vcenter",
})
worksheet.merge_range("A1:F4", "DEEPSEA TECHNOLOGIES", title_format)
worksheet.write("G1", "Projeto\n(Project)", border_format)
worksheet.write("G2", "Cliente\n(Client)", border_format)
worksheet.write("G3", "Número de Projeto\n(Project Number)", border_format)
worksheet.write("G4", "Desenho\n(drawing)",border_format)
worksheet.merge_range("J1:L1", "Número do documento\n(Document Number)", border_format)
worksheet.merge_range("J2:L2", "Revisão\n(Revision)", border_format)
worksheet.merge_range("J3:L3", "Data\n(Date)",border_format)
worksheet.merge_range("J4:L4", "Preparado por\n(Prepared By)", border_format)
worksheet.write("A5", "Linha Nº\n(Line No)", border_format)
worksheet.write("B5","Weld ID\n(ID solda)", border_format)
worksheet.write("C5","Material 1\n(Material 1)", border_format)
worksheet.merge_range("C5:E5", "Material 1\n(Material 1)", border_format)
worksheet.merge_range("F5:H5", "Material 2\n(Material 2)", border_format)
worksheet.write("I5","WPS\n(WPS)", border_format)
worksheet.write("C5","Material 1\n(Material 1)", border_format)
worksheet.merge_range("J5:K5", "WELDER ID\n(ID soldador)", border_format)
worksheet.write("L5", "WELD DATE\n(Data da solda)", border_format)
worksheet.write("M5", "NDE DATE\n(Data NDE)", border_format)
worksheet.write("N5", "VS", border_format)
worksheet.write("O5", "LP/PM", border_format)
worksheet.write("P5", "US/RX", border_format)
worksheet.merge_range("H1:I1", project, border_simple)
worksheet.merge_range("H2:I2", client, border_simple)
worksheet.merge_range("H3:I3", num_project, border_simple)
worksheet.merge_range("H4:I4", drawing, border_simple)
worksheet.merge_range("M1:P1", num_document, border_simple)
worksheet.merge_range("M2:P2", revision, border_simple)
worksheet.merge_range("M3:P3", date, border_simple)
worksheet.merge_range("M4:P4", prepared_by, border_simple)
if len(weld)>0:
    for i in range(0,len(weld)):
        worksheet.write_row(i+5, 1, weld, border_simple)
        worksheet.write_row(i+5, 2, mat1, border_simple)
        worksheet.write_row(i+5, 5, mat2, border_simple)
        worksheet.write_row(i+5, 8, w, border_simple)
        worksheet.write_row(i+5, 9, welder, border_simple)
        worksheet.write_row(i+5, 11, wdate, border_simple)
        worksheet.write_row(i+5, 12, ndate, border_simple)
        worksheet.write_row(i+5, 13, v, border_simple)
        worksheet.write_row(i+5, 14, lp, border_simple)
        worksheet.write_row(i+5, 15, us, border_simple)   
worksheet.merge_range("Q1:V1", "Observação\n(note)", border_simple)
worksheet.merge_range("Q2:V6", note, border_simple)
worksheet.set_column("A:A", 17)
worksheet.set_column("B:B", 17)
worksheet.set_column("G:G", 34)
worksheet.set_column("I:I", 10.5)
worksheet.set_column("J:J", 13.5)
worksheet.set_column("L:L", 25.5)
worksheet.set_column("M:M", 22.5)
worksheet.set_column("A:A", 16.11)

workbook.close()











