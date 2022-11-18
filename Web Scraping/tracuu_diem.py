import csv
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep


#access url
driver = webdriver.Edge()
url = "http://diemthi.hcm.edu.vn/"
driver.get(url)
sleep(1)

#tphcm có hơn 85000 thí sinh, số báo danh có prefix 020 và 5 chữ số tiếp theo gán cho từng thí sinh tham gia
# list_SBD = ["02000100","020000317","02000086"]
list_SBD = []
for i in range(2000001,2003000):
    SBD = '0'+str(i)
    list_SBD.append(SBD)

name_list = []
result_list = []
# list_str = []
remove_SBD = []

#TẠO FILE CSV CÙNG VỚI HEADER
with open('output.csv','w', newline= '',encoding="utf-8") as file_ouput:
    # header = ['Name','Result']
    header = ['Toán', 'Ngữ Văn', 'Vật lí', 'Hóa Học', 'Sinh học', 'KHTN', 'Lịch sử', 'Địa lí', 'GDCD', 'KHXH',
              'Tiếng Anh', 'Tiếng Đức', 'Tiếng Nhật']
    writer = csv.DictWriter(file_ouput, delimiter=',', lineterminator='\n',fieldnames=header)
    writer.writeheader()


#quét SBD lấy tên và kết quả điểm
for SBD_id in list_SBD:
    print(SBD_id)
    id_field = driver.find_element("id","SoBaoDanh")
    id_field.send_keys(SBD_id)
    sleep(1.5)

    xem_diem = driver.find_element("xpath",'//*[@id="loginForm"]/form/div[2]/div/input')
    xem_diem.click()


    page_source = BeautifulSoup(driver.page_source,"html.parser")

    info_div = page_source.find("div", class_ = "container body-content")
    info_label = info_div.find_all('label')
    # print(info_div.find_all('label'))
    label = str(info_label)
    if "Không tìm thấy số báo danh này" in label:
        # list_SBD.remove(SBD_id)
        remove_SBD.append(SBD_id)
        sleep(0.5)
        driver.back()
        sleep(0.5)
        driver.refresh()
        continue
    # print(info_label)
    info_loc = info_div.find_all('td')
    # print(type(info_loc))
    # print(info_loc)
    # name = info_loc[0].get_text().strip()
    # score = info_loc[2].get_text().strip()
    name_result = info_loc[3].get_text().strip()
    score_result = info_loc[5].get_text().strip()
    info_loc[0].get_text().strip()
    # print(score_result)
    # print(f'{name}:{name_result},{score}:{score_result}')
    name_list.append(name_result)
    result_list.append(score_result)

    # write_content(name_result, score_result)
    sleep(0.5)
    driver.back()
    sleep(0.5)
    driver.refresh()

#xóa những SBD không tìm thấy dữ liệu
for id in remove_SBD:
    list_SBD.remove(id)

#tạo list các môn
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
tiengduc = []
tiengnhat = []

#xử lý kết quả chia về các list của từng môn thi
for str in result_list:
    new_str = str.replace(":   ",":").replace("   ",",")
    print(new_str)
    if 'Tiếng Nhật' in new_str:
        if 'KHTN' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Vật lí:","").replace("Hóa học:","").replace("Sinh học:","").replace("KHTN:","").replace("Tiếng Nhật:","")
            newest_str = newest_str.split(',')
            for i in range(6,13):
                newest_str.insert(i,'')
        elif 'KHXH' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Lịch sử:","").replace("Địa lí:","").replace("GDCD:","").replace("KHXH:","").replace("Tiếng Nhật:","")
            newest_str = newest_str.split(',')
            for a in range(6,9):
                newest_str.insert(a,'')
            for i in range(2,6):
                newest_str.insert(i,'')
    elif 'Tiếng Anh' in new_str:
        if 'KHTN' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Vật lí:","").replace("Hóa học:","").replace("Sinh học:","").replace("KHTN:","").replace("Tiếng Anh:","")
            newest_str = newest_str.split(',')
            for a in range(3):
                newest_str.append('')
            for i in range(6,10):
                newest_str.insert(i,'')
        elif 'KHXH' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Lịch sử:","").replace("Địa lí:","").replace("GDCD:","").replace("KHXH:","").replace("Tiếng Anh:","")
            newest_str = newest_str.split(',')
            for b in range(3):
                newest_str.append('')
            for i in range(2,6):
                newest_str.insert(i,'')
    elif 'Tiếng Đức' in new_str:
        if 'KHTN' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Vật lí:","").replace("Hóa học:","").replace("Sinh học:","").replace("KHTN:","").replace("Tiếng Đức:","")
            newest_str = newest_str.split(',')
            for a in range(2):
                newest_str.append('')
            for i in range(6,11):
                newest_str.insert(i,'')
        elif 'KHXH' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Lịch sử:","").replace("Địa lí:","").replace("GDCD:","").replace("KHXH:","").replace("Tiếng Đức:","")
            newest_str = newest_str.split(',')
            for a in range(2):
                newest_str.append('')
            newest_str.insert(6,'')
            for i in range(2,6):
                newest_str.insert(i,'')
    elif 'Tiếng Pháp' in new_str:
        if 'KHTN' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Vật lí:","").replace("Hóa học:","").replace("Sinh học:","").replace("KHTN:","").replace("Tiếng Pháp:","")
            newest_str = newest_str.split(',')
            newest_str.append('')
            for i in range(6,12):
                newest_str.insert(i,'')
        elif 'KHXH' in new_str:
            newest_str = new_str.replace("Toán:","").replace("Ngữ văn:","").replace("Lịch sử:","").replace("Địa lí:","").replace("GDCD:","").replace("KHXH:","").replace("Tiếng Pháp:","")
            newest_str = newest_str.split(',')
            newest_str.append('')
            for a in range(2):
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
    tiengduc.append(newest_str[11])
    tiengnhat.append(newest_str[12])


#XUẤT RA FILE CSV
# df = pd.read_csv("output.csv",names=['Name','Result'])
# diemthi_dict = {'Name':name_list, 'Result':result_list}
diemthi_dict = {'SBD':list_SBD,
                'Name':name_list,
                'Toán':toan,
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
        'Tiếng Đức': tiengduc,
        'Tiếng Nhật': tiengnhat}

# print(diemthi_dict)
df = pd.DataFrame(diemthi_dict)
df.to_csv('output.csv',encoding="utf-8",index=False)

#XUẤT RA FILE EXCEL
new_df = pd.read_csv("output.csv")
print(new_df)
new_df.to_excel("output_excel.xlsx",index=False)

# id_field = driver.find_element("id","SoBaoDanh")
# id_field.send_keys("00222002")
# sleep(1.5)
#
# xem_diem = driver.find_element("xpath",'//*[@id="loginForm"]/form/div[2]/div/input')
# xem_diem.click()
# sleep(3)