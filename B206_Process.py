import openpyxl
import xlsxwriter
import glob
from datetime import date
#מעבר למיקומים שמתעדכנים לפי השנה
current_year=date.today().year-1
current_year=str(current_year)
trueto=  '31.12.'+current_year

create_location1='J:\\מיקום\\' +(current_year)+ '\\בנק ...\\*.txt'
create_location2='J:\\מיקום\\' +(current_year)+ '\\בנק ...\\*.txt'
create_location3='J:\\מיקום\\' +(current_year)+ '\\בנק ...\\*.txt'
create_location4='J:\\מיקום\\' +(current_year)+ '\\בנק ...\\*.txt'
create_location5='J:\\מיקום\\' +(current_year)+ '\\בנק ... \\*.txt'
create_location6='J:\\מיקום\\' +(current_year)+ '\\הבנק ...\\*.txt'
create_location7='J:\\מיקום\\' +(current_year)+ '\\בנק ...\\*.txt'
create_location8='J:\\מיקום\\' +(current_year)+ '\\בנק ...\\*.txt'
create_location9='J:\\מיקום\\' +(current_year)+ '\\בנק.. \\*.txt'
create_location10='J:\\מיקום\\' +(current_year)+ '\\בנק ...\\*.txt'
folders1=[create_location1,create_location2,create_location3,create_location4,create_location5,create_location6,create_location7,create_location8,create_location9,create_location10,]
for i in folders1:
    print(i)
#-----------------------------------------------------------------------
n1_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[0]
    name = file_name.split("\\")
    bankname = name[6]
    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        b_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and line!= ['', '', '\n'] and counter>2:
                line.append(bankname)
                b_list.append(line)
        if b_list==[]:
            b_list.append(bankname)
    n1_list.append(b_list)
#----------------------------------------------------------------------------------------------------------------------------------------------------
n2_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[1]
    name = file_name.split("\\")
    bankname = name[6]
    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        b_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>2:
                line.append(bankname)
                b_list.append(line)
    n2_list.append(b_list)
#------------------------------------------------------------------------------------------נספח ג----------------------------------------------------------
n3_list=[]
b_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[2]
    name = file_name.split("\\")
    bankname = name[6]
    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        c_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>3 and line!=['', '', '', '', '', '', '', '', '', '', '\n']:
                line.append(bankname)
                c_list.append(line)

    n3_list.append(c_list)
#----------------------------------------------------------------------------------------------נספח ד------------------------------------------------------
n4_list=[]
table24=[]
n4_list_coments=[]
n4_list_headers=[]
n4_list_all_lines=[]
header=[]
coment=[]
new_line=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[3]
    name = file_name.split("\\")
    bankname = name[6]
    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        new_line=[]
        all_line=[]
        coment = []
        header = []

        all_line.append(bankname)
        coment.append(bankname)
        new_line.append(bankname)
        header.append(bankname)
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n']and len(line)>2 and line!=['', '', '', '', '', '', '', '', '', '', '\n'] and line!=['', '', '\n'] and counter<17:
                coment.append(line[2])
                all_line.append(line)
                new_line.append(line[1])
                header.append(line[0])
            if counter>19 and line!=['', '', '\n'] and line[0]!='סוג ייעוץ':
                line.append(bankname)
                table24.append(line)
        n4_list_headers.append(header)
        n4_list_coments.append(coment)
        n4_list.append(new_line)
#--------------------------------------------------------------------נספח ה--------------------------------------------------------------------------------
n5_list=[]
d5_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[4]
    name = file_name.split("\\")
    bankname = name[6]

    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        d5_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>2 and line!=['', '', '', '', '', '', '', '', '', '', '\n']:
                    line.append(bankname)
                    d5_list.append(line)
    n5_list.append(d5_list)
#--------------------------------------------------------------------נספח ו---------------------------------------------------------------------------------------
n6_list=[]
d6_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[5]
    name = file_name.split("\\")
    bankname = name[6]
    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        d6_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>3 and line!=['', '', '', '', '', '', '', '', '', '', '\n']:
                    line.append(bankname)
                    d6_list.append(line)
    n6_list.append(d6_list)
#---------------------------------------------------------------------------------נספח ז --------------------------------------------------------------------------
n7_list=[]
d7_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[6]
    name = file_name.split("\\")
    bankname = name[6]

    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        d7_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>2 and line!=['', '', '', '', '', '', '', '', '', '', '\n']:
                    line.append(bankname)
                    d7_list.append(line)
    n7_list.append(d7_list)
#----------------------------------------------------------------נספח ח -------------------------------------------------------------------------------------------
n8_list=[]
n8_list_coments=[]
n8_list_headers=[]
new_line=[]
header=[]
coment=[]

for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[7]
    name = file_name.split("\\")
    bankname = name[6]

    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        new_line=[]
        new_line.append(bankname)

        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and line!=['', '', '', '', '', '', '', '', '', '', '\n'] and line!=['', '', '\n']:

                new_line.append(line[1])
        n8_list.append(new_line)
#----------------------------------------------------------------נספח ט -------------------------------------------------------------------------------------------
n9_list=[]
d9_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[8]
    name = file_name.split("\\")
    bankname = name[6]

    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        d9_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>2 and line!=['', '', '', '', '', '', '', '', '', '', '\n']:

                    line.append(bankname)
                    d9_list.append(line)

    n9_list.append(d9_list)
#----------------------------------------------------------------נספח י-------------------------------------------------------------------------------------------
n10_list=[]
d10_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[9]
    name = file_name.split("\\")
    bankname = name[6]

    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        d10_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>2 and line!=['', '', '', '', '', '', '', '', '', '', '\n'] and line!=['', '', '', '', '', '\n']:
                    line.append(bankname)
                    d10_list.append(line)

    n10_list.append(d10_list)
#----------------------------------------------------------------נספח יא-------------------------------------------------------------------------------------------
n11_list=[]
d11_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[10]
    name = file_name.split("\\")
    bankname = name[6]

    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        d11_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>2 and line!=['', '', '', '', '', '', '', '', '', '', '\n'] and line!=['', '', '', '', '', '\n'] and line!=['', '', '', '', '', '', '\n']:
                    line.append(bankname)
                    d11_list.append(line)
    n11_list.append(d11_list)
#----------------------------------------------------------------נספח יב-------------------------------------------------------------------------------------------
n12_list=[]
d12_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[11]
    name = file_name.split("\\")
    bankname = name[6]

    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        d11_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>2 and line!=['', '', '', '', '', '', '', '', '', '', '\n'] and line!=['', '', '', '', '', '\n'] and line!=['', '', '', '', '', '', '\n'] and line!=['', '\n']:
                line.append(bankname)
                d11_list.append("l")
                d12_list.append(line)
    n12_list.append(d12_list)
#___________________________________________________________________נספח יג _____________________________________________
n13_list=[]
d13_list=[]
for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[12]
    name = file_name.split("\\")
    bankname = name[6]

    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        d13_list=[]
        counter=0
        for line in f:
            counter+=1
            line = line.split("\t")
            if line!=['', '', '', '', '\n'] and counter>2 and line!=['', '', '', '', '', '', '', '', '', '', '\n'] and line!=['', '', '', '', '', '\n'] and line!=['', '', '', '', '', '', '\n'] and line!=['', '\n']:
                    line.append(bankname)
                    d13_list.append(line)

    n13_list.append(d13_list)
#----------------------------------------------------------------------------------------------------שאלון ראשי---------------------------------------------------
n14_list=[]
d14_list=[]
segments=['1','2','3','4','4א','5','5.1','6','6.1','7','7.1','7.2','7.3','7.4','8','8.1','9', '10','11','11.1','12','12א','13','13א' , '13ב','13ג','14','14א','14.1א','15','16','17','18','19','20','21','22', '23','24','25','25א'  ,'25ב' ,'26','27','28','28.1','28.2','28.3','29']

for bank in folders1:
    folder_name = glob.glob(bank)
    file_name=folder_name[13]
    name = file_name.split("\\")
    bankname = name[6]
    with open(file_name, encoding='windows-1255', errors='ignore') as f:
        d14_list=[]
        counter=0
        count2=1
        for line in f:

            counter+=1
            line = line.split("\t")
            if  bankname!='בנק ...' and counter > 6 and line[0] in segments:
                    count2+=1

                    line.append(bankname)

                    line.append(count2)
                    n=len(line)
                    temp = line[0]
                    line[0] = line[n - 1]
                    line[n - 1] = temp
                    n14_list.append(line)
                    n14_list.sort()

#--------------------------------------------------------------------------------------------------------------------------------------------------

#---------------------------------------------------------#פעולות של ייצוא לאקסל ועיצוב הקובץ---------------------------------------------------
name= 'שם קובץ'+current_year+'.xlsx'
name= 'J:\\מיקום\\' + name
outworkbook=xlsxwriter.Workbook(name)
#--------------------------------------------------------------פותח גליונות ומגדיר מימין לשמאל------------------------------------
worksheet14=outworkbook.add_worksheet('')
worksheet1=outworkbook.add_worksheet(' נספח א')
worksheet2=outworkbook.add_worksheet(' נספח ב')
worksheet3=outworkbook.add_worksheet(' נספח ג')
worksheet4=outworkbook.add_worksheet(' נספח ד')
worksheet5=outworkbook.add_worksheet(' נספח ה')
worksheet6=outworkbook.add_worksheet(' נספח ו')
worksheet7=outworkbook.add_worksheet(' נספח ז')
worksheet8=outworkbook.add_worksheet(' נספח ח')
worksheet9=outworkbook.add_worksheet(' נספח ט')
worksheet10=outworkbook.add_worksheet(' נספח י')
worksheet11=outworkbook.add_worksheet(' נספח יא')
worksheet12=outworkbook.add_worksheet(' נספח יב')
worksheet13=outworkbook.add_worksheet(' נספח יג')

worksheet1.right_to_left()
worksheet2.right_to_left()
worksheet3.right_to_left()
worksheet4.right_to_left()
worksheet5.right_to_left()
worksheet6.right_to_left()
worksheet7.right_to_left()
worksheet8.right_to_left()
worksheet9.right_to_left()
worksheet10.right_to_left()
worksheet11.right_to_left()
worksheet12.right_to_left()
worksheet13.right_to_left()
worksheet14.right_to_left()
#--------------------------------------------------------------פותח גליונות ומגדיר מימין לשמאל------------------------------------
header= "כותרת'"
merge_format = outworkbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'underline':True,
    'valign': 'vcenter'})
bold=outworkbook.add_format({'bold':True})
bold.set_text_wrap()
cell_format=outworkbook.add_format()
cell_format.set_text_wrap()
worksheet1.set_row(0,60)
worksheet1.set_row(1,40)
worksheet1.set_column('A:A',12)
worksheet1.set_column('B:D',40)
worksheet1.merge_range('A1:D1',header,merge_format)

#כותרות לטבלה

counter=2
for i in range(len(n1_list)):
    for j in range(len(n1_list[i])):

        single_line=n1_list[i][j]

        bank_name = single_line[len(single_line) - 1]
        company=single_line[0]
        company_held = single_line[1]
        company_charachter = single_line[2]
        company_precent = single_line[3]
        colums=1
        worksheet1.write(counter,0,bank_name,cell_format)
        try:
            for e in range(len(single_line)-1):
                try:
                    number = single_line[e]

                    single_line[e] = float(number.replace('"', '').replace(',', ''))
                except:
                    single_line[e] = single_line[e]
                worksheet1.write(counter,colums,single_line[e],cell_format)
                colums+=1
            counter+=1
        except:
            worksheet1.write(counter, 0, single_line,cell_format)
            counter+=1
tablelimit = counter
tablelimit_new = 'A2:D' + str(tablelimit)
worksheet1.add_table(tablelimit_new)
worksheet1.write("A2"," טקסט רצוי",cell_format)
worksheet1.write("B2","טקסט רצוי",cell_format)
worksheet1.write("C2","טקסט רצוי",cell_format)
worksheet1.write("D2","טקסט רצוי",cell_format)

#    --------------------------------------------------------------גליון 2--------------------------------
counter=2
for i in range(len(n2_list)):
    for j in range(len(n2_list[i])):

        single_line=n2_list[i][j]
        counter += 1
        if len(single_line)>3:
            bank_name = single_line[len(single_line) - 1]
            company=single_line[0]
            company_held = single_line[1]
            company_charachter = single_line[2]
            company_precent = single_line[3]
            note = single_line[4]
            worksheet2.write(counter,0,bank_name,cell_format)
            worksheet2.write(counter, 1, company, cell_format)
            worksheet2.write(counter, 2, company_held, cell_format)
            worksheet2.write(counter,3, company_charachter, cell_format)
            worksheet2.write(counter, 4, company_precent, cell_format)
            worksheet2.write(counter, 5, note, cell_format)
tablelimit = counter + 2
tablelimit_new = 'A3:F' + str(tablelimit)
worksheet2.add_table(tablelimit_new)

header='כותרת'
worksheet2.merge_range('A1:F2',header,merge_format)
worksheet2.write("A3","טקסט רצוי",cell_format)
worksheet2.write("B3","טקסט רצוי",cell_format)
worksheet2.write("C3","טקסט רצוי",cell_format)
worksheet2.write("D3","טקסט רצוי)",cell_format)
worksheet2.write("D3","טקסט רצוי)",cell_format)
worksheet2.write("E3","טקסט רצוי)",cell_format)
worksheet2.write("F3","טקסט רצוי",cell_format)

#    --------------------------------------------------------------גליון 3--------------------------------
counter=4
for i in range(len(n3_list)):
    for j in range(len(n3_list[i])):

        single_line=n3_list[i][j]
        if len(single_line)>3:
            bank_name = single_line[len(single_line) - 1]
            company=single_line[0]
            company_held = single_line[1]
            company_charachter = single_line[2]
            company_precent = single_line[3]
            note = single_line[4]
            colums=1
            worksheet3.write(counter,0,bank_name,cell_format)
            for e in range(len(single_line)-1):
                try:
                    number = single_line[e]

                    single_line[e] = float(number.replace('"', '').replace(',', '').replace('.', ''))
                except:
                    single_line[e]=single_line[e]
                worksheet3.write(counter,colums,single_line[e],cell_format)
                colums+=1
            counter+=1

tablelimit = counter + 2
tablelimit_new = 'A4:L' + str(tablelimit)
worksheet3.add_table(tablelimit_new)
header='טקסט רצוי'
worksheet3.merge_range('A1:L2',header,merge_format)
worksheet3.write("A4",'',cell_format)
worksheet3.write("B4",'',cell_format)

worksheet3.write("C3","טקסט רצוי",cell_format)
worksheet3.write("D3","טקסט רצוי",cell_format)
worksheet3.write("E3","טקסט רצוי",cell_format)
worksheet3.write("F3","טקסט רצוי",cell_format)
worksheet3.write("G3","טקסט רצוי",cell_format)
worksheet3.write("H3","טקסט רצוי",cell_format)
worksheet3.write("I3","טקסט רצוי",cell_format)
worksheet3.write("J3","טקסט רצוי",cell_format)
worksheet3.write("K3","טקסט רצוי",cell_format)
worksheet3.write("L3","טקסט רצוי",cell_format)
worksheet3.write("A4","טקסט רצוי",cell_format)
worksheet3.write("B4","טקסט רצוי",cell_format)
worksheet3.write("C4","טקסט רצוי",cell_format)
worksheet3.write("D4","טקסט רצוי",cell_format)
worksheet3.write("E4","טקסט רצוי",cell_format)
worksheet3.write("F4","טקסט רצוי",cell_format)
worksheet3.write("G4","טקסט רצוי",cell_format)
worksheet3.write("H4","טקסט רצוי",cell_format)
worksheet3.write("I4","טקסט רצוי",cell_format)
worksheet3.write("J4","טקסט רצוי",cell_format)
worksheet3.write("K4","טקסט רצוי",cell_format)
worksheet3.write("L4","טקסט רצוי",cell_format)

#    --------------------------------------------------------------גליון 4------------------------------
header='כותרת'
colums=2
counter=3
tablelimit_new = 'A4:L18'
worksheet4.add_table(tablelimit_new)
for i in range(len(n4_list)):
    for j in range(len(n4_list[i])):
        head= n4_list[i][2]+" " +n4_list[i][0]


        if n4_list[i][j]!='':
            try:
                number = n4_list[i][j]

                n4_list[i][j] = float(number.replace('"', '').replace(',', '').replace('.', ''))
            except:
                n4_list[i][j] = n4_list[i][j]

            worksheet4.write(j+1, i+1, n4_list[i][j], cell_format)
            worksheet4.write(j+1, 11, n4_list_coments[i][j], cell_format)
            worksheet4.write(j+1, 0, n4_list_headers[i][j], cell_format)
    worksheet4.write(3, i+1, head, cell_format)
counter = 21
for z in range(len(table24)-1):
    worksheet4.write(counter, 0, table24[z][len(table24[z])-1], cell_format)
    worksheet4.write(counter, 1, table24[z][0], cell_format)
    try:
        number = table24[z][1]

        table24[z][1] = float(number.replace('"', '').replace(',', ''))
    except:
        table24[z][1] = table24[z][1]
    worksheet4.write(counter, 2, table24[z][1], cell_format)
    worksheet4.write(counter, 3, table24[z][2], cell_format)

    counter+=1

tablelimit = counter
tablelimit_new = 'A21:D' + str(tablelimit)
worksheet4.add_table(tablelimit_new)
header='כותרת'
worksheet4.merge_range('A20:D20',header,merge_format)
worksheet4.write("A21",'טקסט רצוי',cell_format)
worksheet4.write("B21",'טקסט רצוי',cell_format)
worksheet4.write("C21",'טקסט רצוי',cell_format)
worksheet4.write("D21",'טקסט רצוי',cell_format)


tablelimit = counter+2
worksheet4.merge_range('A1:L3',header,merge_format)
#    --------------------------------------------------------------גליון 5-----------------------------
header='כותרת'
counter=3

for i in range(len(n5_list)):
    for j in range(len(n5_list[i])):
        single_line=n5_list[i][j]
        if len(single_line)>3:
            bank_name = single_line[len(single_line) - 1]
            company=single_line[0]
            company_held = single_line[1]
            company_charachter = single_line[2]
            company_precent = single_line[3]
            note = single_line[4]
            colums=1
            worksheet5.write(counter,0,bank_name,cell_format)
            for e in range(len(single_line)-1):
                try:
                    number = single_line[e]

                    single_line[e] = float(number.replace('"', '').replace(',', '').replace('.', ''))
                except:
                    single_line[e] = single_line[e]
                worksheet5.write(counter,colums,single_line[e],cell_format)
                colums+=1
            counter+=1
tablelimit = counter+2
tablelimit_new = 'A3:F' + str(tablelimit)

worksheet5.add_table(tablelimit_new)
worksheet5.merge_range('A1:F2',header,merge_format)
worksheet5.write("A3","טקסט רצוי ",cell_format)
worksheet5.write("B3","טקסט רצוי",cell_format)
worksheet5.write("C3","טקסט רצוי",cell_format)
worksheet5.write("D3","טקסט רצוי",cell_format)
worksheet5.write("E3","טקסט רצוי",cell_format)
worksheet5.write("F3","טקסט רצוי",cell_format)
#    --------------------------------------------------------------גליון 6-----------------------------
header='כותרת'+current_year



counter=3
for i in range(len(n6_list)):
    for j in range(len(n6_list[i])-1):

        single_line=n6_list[i][j]
        if len(single_line)>3:
            bank_name = single_line[len(single_line) - 1]
            company=single_line[1]
            company_held = single_line[2]
            company_charachter = single_line[3]
            company_precent = single_line[4]
            note = single_line[5]
            colums=1
            worksheet6.write(counter,0,bank_name,cell_format)
            for e in range(len(single_line)-1):
                try:
                    number = single_line[e]

                    single_line[e] = float(number.replace('"', '').replace(',', '').replace('.', ''))
                except:
                    single_line[e]=single_line[e]
                worksheet6.write(counter,colums,single_line[e],cell_format)
                colums+=1
            counter+=1
tablelimit = counter
tablelimit_new = 'A4:F' + str(tablelimit)

worksheet6.add_table(tablelimit_new)
worksheet6.merge_range('A1:F2',header,merge_format)
worksheet6.write("C3","טקסט רצוי ",cell_format)
worksheet6.write("D3","טקסט רצוי ",cell_format)
worksheet6.write("A4","טקסט רצוי",cell_format)
worksheet6.write("B4","טקסט רצוי",cell_format)
worksheet6.write("C4","טקסט רצוי "+trueto)
worksheet6.write("D4",'טקסט רצוי '+ current_year+'-טקסט רצוי'+current_year)
worksheet6.write("E4","טקסט רצוי "+current_year)
worksheet6.write("F4","טקסט רצוי",cell_format)

#    --------------------------------------------------------------גליון 7-----------------------------
header='כותרת

counter=3
for i in range(len(n7_list)):
    for j in range(len(n7_list[i])):

        single_line=n7_list[i][j]
        if len(single_line)>3:
            bank_name = single_line[len(single_line) - 1]
            company=single_line[0]
            company_held = single_line[1]
            company_charachter = single_line[2]
            company_precent = single_line[3]
            note = single_line[4]
            colums=1
            worksheet7.write(counter,0,bank_name,cell_format)
            for e in range(len(single_line)-1):
                try:
                    if e!=2:
                        number = single_line[e]

                    single_line[e] = float(number.replace('"', '').replace(',', '').replace('.', ''))
                except:
                    single_line[e]=single_line[e]

                worksheet7.write(counter,colums,single_line[e],cell_format)
                worksheet7.write(counter, 3, single_line[2],cell_format)
                colums+=1
            counter+=1
tablelimit = counter+2
tablelimit_new = 'A3:F' + str(tablelimit)

worksheet7.add_table(tablelimit_new)
worksheet7.merge_range('A1:F1',header,merge_format)
worksheet7.write("A3","טקסט רצוי ",cell_format)
worksheet7.write("B3","טקסט רצוי ",cell_format)
worksheet7.write("D3","טקסט רצוי",cell_format)
worksheet7.write("C3","טקסט רצוי",cell_format)
worksheet7.write("E3",'טקסט רצוי',cell_format)
worksheet7.write("F3",'טקסט רצוי' ,cell_format)

#    --------------------------------------------------------------גליון 8-----------------------------
header='כותרת' +current_year
colums=1
counter=2
tablelimit_new = 'A3:K7'
worksheet8.add_table(tablelimit_new)
for i in range(len(n8_list)):
    for j in range(len(n8_list[i])):
        head= n8_list[i][2]+"-" +n4_list[i][0]
        number=0
        try:
            number = n8_list[i][j]

            n8_list[i][j]=float(number.replace('"','').replace(',','').replace('.',''))
        except:
            number = n8_list[i][j]




        worksheet8.write(j, i + 1, n8_list[i][j], cell_format)
    worksheet8.write(2, 0, 'טקסט רצוי', cell_format)
    worksheet8.write(3, 0, 'טקסט רצוי '+current_year,cell_format)
    worksheet8.write(4, 0,'טקסט רצוי '+current_year, cell_format)
    worksheet8.write(5,0, 'טקסט רצוי ' +current_year, cell_format)
    worksheet8.write(6,0, ' טקסט רצוי ' +current_year,cell_format)
    worksheet8.write(2, i+1, head, cell_format)

tablelimit = counter+2
worksheet8.merge_range('A1:K2',header,merge_format)

#    --------------------------------------------------------------גליון 9-----------------------------
header='כותרת'

counter=3
for i in range(len(n9_list)):
    for j in range(len(n9_list[i])):

        single_line=n9_list[i][j]
        if len(single_line)>3:
            bank_name = single_line[len(single_line) - 1]
            company=single_line[0]
            company_held = single_line[1]
            company_charachter = single_line[2]
            company_precent = single_line[3]
            note = single_line[4]
            colums=1
            worksheet9.write(counter,0,bank_name,cell_format)
            for e in range(len(single_line)-1):
                worksheet9.write(counter,colums,single_line[e],cell_format)
                colums+=1
            counter+=1
tablelimit = counter+2
tablelimit_new = 'A3:F' + str(tablelimit)

worksheet9.add_table(tablelimit_new)
worksheet9.merge_range('A1:F1',header,merge_format)
worksheet9.write("A3","טקסט רצוי ",cell_format)
worksheet9.write("B3","טקסט רצוי",cell_format)
worksheet9.write("C3","טקסט רצוי",cell_format)
worksheet9.write("D3","טקסט רצוי",cell_format)
worksheet9.write("E3",'טקסט רצוי',cell_format)
worksheet9.write("F3",'טקסט רצוי' ,cell_format)

#---------------------------------------------------------גליון 10----------------------------------------------------------------
header='כותרת'

counter=2
for i in range(len(n10_list)):
    for j in range(len(n10_list[i])):

        single_line=n10_list[i][j]
        if len(single_line)>3:
            bank_name = single_line[len(single_line) - 1]
            company=single_line[0]
            company_held = single_line[1]
            company_charachter = single_line[2]
            company_precent = single_line[3]
            note = single_line[4]
            colums=1
            worksheet10.write(counter,0,bank_name,cell_format)
            for e in range(len(single_line)-1):
                worksheet10.write(counter,colums,single_line[e],cell_format)
                colums+=1
            counter+=1
tablelimit = counter+2
tablelimit_new = 'A2:G' + str(tablelimit)

worksheet10.add_table(tablelimit_new)
worksheet10.merge_range('A1:G1',header,merge_format)
worksheet10.write("A2","טקסט רצוי ",cell_format)
worksheet10.write("B2","טקסט רצוי",cell_format)
worksheet10.write("D2","טקסט רצוי",cell_format)
worksheet10.write("C2","טקסט רצוי",cell_format)
worksheet10.write("E2",'טקסט רצוי',cell_format)
worksheet10.write("F2",'טקסט רצוי',cell_format)
worksheet10.write("G2",'טקסט רצוי' ,cell_format)


#---------------------------------------------------------גליון 11----------------------------------------------------------------
header=' כותרת'

counter=2
for i in range(len(n11_list)):
    for j in range(len(n11_list[i])):

        single_line=n11_list[i][j]
        if len(single_line)>3:
            bank_name = single_line[len(single_line) - 1]
            company=single_line[0]
            company_held = single_line[1]
            company_charachter = single_line[2]
            company_precent = single_line[3]
            note = single_line[4]
            colums=1
            worksheet11.write(counter,0,bank_name,cell_format)
            for e in range(len(single_line)-1):
                worksheet11.write(counter,colums,single_line[e],cell_format)
                colums+=1
            counter+=1
tablelimit = counter+2
tablelimit_new = 'A2:I' + str(tablelimit)

worksheet11.add_table(tablelimit_new)
worksheet11.merge_range('A1:I1',header,merge_format)
worksheet11.write("A2","טקסט רצוי ",cell_format)
worksheet11.write("B2","טקסט רצוי ",cell_format)
worksheet11.write("D2","טקסט רצוי ",cell_format)
worksheet11.write("C2","טקסט רצוי",cell_format)
worksheet11.write("E2",'טקסט רצוי',cell_format)
worksheet11.write("F2",'טקסט רצוי',cell_format)
worksheet11.write("G2",'טקסט רצוי' ,cell_format)
worksheet11.write("H2",'טקסט רצוי' ,cell_format)
worksheet11.write("I2",'טקסט רצוי' ,cell_format)

#---------------------------------------------------------גליון 12----------------------------------------------------------------
header='כותרת'
counter=3
for i in range(len(n12_list)):
    for j in range(len(n12_list[i])):

        single_line=n12_list[i][j]
        if len(single_line)==1:
            worksheet12.write(counter, 0, bank_name, cell_format)
        if len(single_line)>1 :
            bank_name = single_line[len(single_line) - 1]
            company=single_line[0]
            company_held = single_line[1]
            colums=1
            worksheet12.write(counter,0,bank_name,cell_format)
            for e in range(len(single_line)-1):
                worksheet12.write(counter,colums,single_line[e],cell_format)
                colums+=1
            counter+=1
tablelimit = counter
tablelimit_new = 'A3:C' + str(tablelimit)

worksheet12.add_table(tablelimit_new)
worksheet12.merge_range('A1:C2',header,merge_format)
worksheet12.write("A3",'טקסט רצוי' ,cell_format)
worksheet12.write("B3",'טקסט רצוי' ,cell_format)
worksheet12.write("C3",'טקסט רצוי' ,cell_format)




#___________________________________________________________________________________________גליון 13____________________________
header='כותרת'+' ' +current_year

counter=3
for i in range(len(n13_list)):
    for j in range(len(n13_list[i])):

        single_line=n13_list[i][j]
        if len(single_line)>1:

            bank_name = single_line[len(single_line) - 1]
            company=single_line[0]
            company_held = single_line[1]


            colums=1
            worksheet13.write(counter,0,bank_name,cell_format)
            for e in range(len(single_line)-1):
                worksheet13.write(counter,colums,single_line[e],cell_format)
                colums+=1
            counter+=1
tablelimit = counter
tablelimit_new = 'A3:C' + str(tablelimit)

worksheet13.add_table(tablelimit_new)
worksheet13.merge_range('A1:C2', header, merge_format)
worksheet13.write("A3", 'טקסט רצוי', cell_format)
worksheet13.write("B3", 'טקסט רצוי', cell_format)
worksheet13.write("C3", 'טקסט רצוי', cell_format)
#    --------------------------------------------------------------גליון שאלון ראשי-----------------------------
header='טקסט רצוי'+' '+current_year
worksheet14.merge_range('A1:G1',header,merge_format)
header='טקסט רצוי'+' '+trueto+ 'טקסט רצוי'
worksheet14.merge_range('A2:G2',header,merge_format)

count=3

for i in range(len(n14_list)):

        worksheet14.write(count, 2, n14_list[i][2], cell_format)
        worksheet14.write(count, 3, n14_list[i][len(n14_list[i]) - 2], cell_format)
        worksheet14.write(count, 0, n14_list[i][len(n14_list[i]) - 1], cell_format)
        worksheet14.write(count, 1, n14_list[i][1], cell_format)
        try:
            number = n14_list[i][3]

            n14_list[i][3]=float(number.replace('"','').replace(',','').replace('.',''))


        except:
            n14_list[i][3]=n14_list[i][3]
            worksheet14.write(count, 4, n14_list[i][2], cell_format)
        worksheet14.write(count, 4, n14_list[i][3], cell_format)
        worksheet14.write(count, 6, n14_list[i][5], cell_format)
        if n14_list[i][4]=='כן' or n14_list[i][4]=='לא' or n14_list[i][4]=='':
            worksheet14.write(count, 5, n14_list[i][4], cell_format)

        count+=1


tablelimit = count
tablelimit_new = 'A3:G' + str(tablelimit)
worksheet14.add_table(tablelimit_new)
worksheet14.write(2, 0, 'טקסט רצוי', cell_format)
worksheet14.write(2, 1, 'טקסט רצוי', cell_format)
worksheet14.write(2, 2, 'טקסט רצוי', cell_format)
worksheet14.write(2, 3, 'טקסט רצוי', cell_format)
worksheet14.write(2, 4, 'טקסט רצוי', cell_format)
worksheet14.write(2, 5, 'טקסט רצוי', cell_format)
worksheet14.write(2, 6, 'טקסט רצוי', cell_format)

#מייצא קובץ לא לגעת:-------------------------------------
worksheet1=openpyxl.load_workbook(name)
worksheet1=worksheet1.active
outworkbook.close()
print(name)
