import pandas as pd
import csv
str1 = "Toán:   4.40   Ngữ văn:   4.50   Vật lí:   3.00   Hóa học:   2.50   Sinh học:   3.50   KHTN: 3   Tiếng Anh:   4.60"
str2 = "Toán:   8.60   Ngữ văn:   6.00   Vật lí:   8.25   Hóa học:   9.00   Sinh học:   5.75   KHTN: 7.67   Tiếng Anh:   8.80"
str3 = "Toán:   8.20   Ngữ văn:   7.50   Lịch sử:   7.25   Địa lí:   7.25   GDCD:   8.25   KHXH: 7.58   Tiếng Anh:   9.20"
str4 = "Toán:   7.20   Ngữ văn:   6.50   Vật lí:   7.25   Hóa học:   5.00   Sinh học:   5.25   KHTN: 5.83   Tiếng Nhật:   6.40"
str5 = "Toán:   8.80   Ngữ văn:   6.75   Vật lí:   5.00   Hóa học:   8.50   Sinh học:   8.25   KHTN: 7.25"
str6 = "Toán:   7.20   Ngữ văn:   6.50   Vật lí:   7.25   Hóa học:   5.00   Sinh học:   5.25   KHTN: 5.83   Tiếng Nhật:   6.40"
list_str = [str1, str2,str3,str4,str5]

# def diemthi_process(list_str):
header = ['Toán','Ngữ Văn','Vật lí','Hóa Học','Sinh học','KHTN','Lịch sử','Địa lí','GDCD','KHXH','Tiếng Anh','Tiếng Pháp','Tiếng Nhật']
toan = []
van = []
vatli = []
hoahoc = []
sinhhoc = []
KHTN = []
tienganh = []
lichsu = []
diali = []
GDCD = []
KHXH = []
tiengphap = []
tiengnhat = []

for str in list_str:
    new_str = str.replace(":   ",":").replace("   ",",")
    print(new_str)
    if 'Tiếng Nhật' in new_str:
        if 'KHTN' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Vật lí:","").replace("Hóa học:","").replace("Sinh học:","").replace("KHTN:","").replace("Tiếng Nhật:","")
            newest_str = newest_str.split(',')
            for i in range(6,12):
                newest_str.insert(a,'')
        elif 'KHXH' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Lịch sử:","").replace("Địa lí:","").replace("GDCD:","").replace("KHXH:","").replace("Tiếng Nhật:","")
            newest_str = newest_str.split(',')
            for a in range(6,8):
                newest_str.insert(a,'')
            for i in range(2,6):
                newest_str.insert(i,'')
    elif 'Tiếng Anh' in new_str:
        if 'KHTN' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Vật lí:","").replace("Hóa học:","").replace("Sinh học:","").replace("KHTN:","").replace("Tiếng Anh:","")
            newest_str = newest_str.split(',')
            for a in range(2):
                newest_str.append('')
            for i in range(6,10):
                newest_str.insert(i,'')
        elif 'KHXH' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Lịch sử:","").replace("Địa lí:","").replace("GDCD:","").replace("KHXH:","").replace("Tiếng Anh:","")
            newest_str = newest_str.split(',')
            for b in range(2):
                newest_str.append('')
            for i in range(2,6):
                newest_str.insert(i,'')
    elif 'Tiếng Pháp' in new_str:
        if 'KHTN' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Vật lí:","").replace("Hóa học:","").replace("Sinh học:","").replace("KHTN:","").replace("Tiếng Pháp:","")
            newest_str = newest_str.split(',')
            newest_str.append('')
            for i in range(6,11):
                newest_str.insert(i,'')
        elif 'KHXH' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Lịch sử:","").replace("Địa lí:","").replace("GDCD:","").replace("KHXH:","").replace("Tiếng Pháp:","")
            newest_str = newest_str.split(',')
            newest_str.append('')
            newest_str.insert(6,'')
            for i in range(2,6):
                newest_str.insert(i,'')
    else:
        if 'KHTN' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Vật lí:","").replace("Hóa học:","").replace("Sinh học:","").replace("KHTN:","")
            newest_str = newest_str.split(',')
            newest_str.append('')
            for i in range(6,13):
                newest_str.insert(i,'')
        elif 'KHXH' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Lịch sử:","").replace("Địa lí:","").replace("GDCD:","").replace("KHXH:","")
            newest_str = newest_str.split(',')
            for a in range(3):
                newest_str.append('')
            for i in range(2,6):
                newest_str.insert(i,'')

    toan.append(newest_str[0])
    van.append(newest_str[1])
    vatli.append(newest_str[2])
    hoahoc.append(newest_str[3])
    sinhhoc.append(newest_str[4])
    KHTN.append(newest_str[5])
    lichsu.append(newest_str[6])
    diali.append(newest_str[7])
    GDCD.append(newest_str[8])
    KHXH.append(newest_str[9])
    tienganh.append(newest_str[10])
    tiengphap.append(newest_str[11])
    tiengnhat.append(newest_str[12])
# return toan, van, vatli, hoahoc, KHTN, tienganh, lichsu, diali, GDCD, KHXH

with open('diemthi.csv','w', newline= '',encoding="utf-8") as diemthi_ouput:
    # header = ['Toán','Ngữ Văn','Vật lí','Hóa Học','Sinh học','KHTN','Tiếng Anh']
    writer = csv.DictWriter(diemthi_ouput, delimiter=',', lineterminator='\n',fieldnames=header)
    writer.writeheader()
    # writer.writerow({[header]:[newest_str1]})
df = pd.read_csv("diemthi.csv")
data = {'Toán':toan,
        'Ngữ Văn':van,
        'Vật lí':vatli,
        'Hóa Học':hoahoc,
        'Sinh học':sinhhoc,
        'KHTN':KHTN,
        'Lịch sử':lichsu,
        'Địa lí':diali,
        'GDCD':GDCD,
        'KHXH':KHXH,
        'Tiếng Anh':tienganh,
        'Tiếng Pháp':tiengphap,
        'Tiếng Nhật':tiengnhat,}
df = pd.DataFrame(data)
print(df)
df.to_csv('diemthi.csv',index=False)

new_df = pd.read_csv("diemthi.csv")
print(new_df)
new_df.to_excel("diemthi_excel.xlsx",index=False)
