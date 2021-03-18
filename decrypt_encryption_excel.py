import win32com.client
import os
# 加密前，excel没有密码
def encryption_excel(folder_path,password):
    xlApp = win32com.client.Dispatch("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    file_name_li=os.listdir(folder_path)
    for file_name in file_name_li:
        file_path=os.path.join(folder_path,file_name)
        xw=xlApp.Workbooks.Open(file_path,False,False,None,"")
        xw.SaveAs(file_path,None,password,"")
    xlApp.Quit()
# 解密
def decrypt_excel(folder_path,password):
    xlApp = win32com.client.Dispatch("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    file_name_li = os.listdir(folder_path)
    for file_name in file_name_li:
        file_path = os.path.join(folder_path, file_name)
        xw = xlApp.Workbooks.Open(file_path, False, False, None, password)
        xw.SaveAs(file_path, None, "", "")
    xlApp.Quit()

if __name__ == '__main__':
    password="Password123"
    with open("excel_config.txt","r",encoding="utf8") as f:
        folder_path=f.readline()
        folder_path=folder_path.strip("\n").split("路径:")[1]
        method=f.readline()
        method=method.strip("\n").split("方法:")[1]
    print(method)
    if method=="解密":
       decrypt_excel(folder_path,password)
    elif method=="加密":
        encryption_excel(folder_path,password)
    with open("result.txt","w",encoding="utf8")as f:
        f.write(f"{folder_path}内excel文件{method}成功")
