import xlwings as xw
import os
import string



def getColumnName(columnIndex):
    """
    获取列索引获取列名,例如输入8，返回H
    :param columnIndex:列索引
    :return:
    """
    ret = ''
    ci = columnIndex - 1
    index = ci // 26
    if index > 0:
        ret += getColumnName(index)
    ret += string.ascii_uppercase[ci % 26]
    return ret




def copy_func(sheet_name,from_wb,tar_wb,header,is_mix,tar_wb_sheet_name_list):
    try:
        from_sh = from_wb.sheets[sheet_name]
    except IndexError:
        return
    sheet_name_info = from_sh.name
    if is_mix:
        tar_sh = tar_wb.sheets[0]
    else:
        if sheet_name_info not in tar_wb_sheet_name_list:
            tar_wb.sheets.add(sheet_name_info)
            tar_wb_sheet_name_list.append(sheet_name_info)
        tar_sh = tar_wb.sheets[sheet_name_info]
    from_sh_range_api=from_sh.range(f"A{header+1}").expand("table").api
    last_cell_row=tar_sh.used_range.last_cell.row if tar_sh.used_range.last_cell.value==None else tar_sh.used_range.last_cell.row+1
    from_sh_range_api.Copy(tar_sh.range("A"+str(last_cell_row)).expand("table").api)


def copy_excel(folder_path,header=1,sheet_name=0,is_mix=False):
    """
    功能：复制一个文件夹内所有的Excel文件的数据，汇总到一个新的文件夹中，
    :param folder_path:所需要的复制的Excel文件
    :param header:默认值为1，第一行为表头，0为不存在表头，复制时不复制表头
    :param sheet_name:可以为数字，字符串和列表，数字时为Sheet索引，字符串时为Sheet名称，列表时为两者混合
    :is_mix:是否混合，默认是False，混合模式下会将所有数据存放到一个sheet中，不混合则将对应的sheet页进行汇总
    :return:
    """
    new_file_path=os.path.join(folder_path,"result.xlsx")
    excel_file_path_list=os.listdir(folder_path)
    app = xw.App(visible=False, add_book=False)
    tar_wb = app.books.add()
    tar_wb.save(new_file_path)
    tar_wb_sheet_name_list=[]
    for i in range(0, tar_wb.sheets.count):
        tar_wb_sheet_name_list.append(tar_wb.sheets[i].name)
    for excel_file_path in excel_file_path_list:
        wb = app.books.open(os.path.join(folder_path,excel_file_path))
        if isinstance(sheet_name,list):
            for sheet_name_info in sheet_name:
                copy_func(sheet_name_info,from_wb=wb,tar_wb=tar_wb,header=header,is_mix=is_mix,tar_wb_sheet_name_list=tar_wb_sheet_name_list)
        else:
            copy_func(sheet_name, from_wb=wb, tar_wb=tar_wb, header=header, is_mix=is_mix,tar_wb_sheet_name_list=tar_wb_sheet_name_list)
        wb.close()
    tar_wb.save(new_file_path)
    app.quit()

if __name__ == '__main__':

    folder_path=r"C:\Users\86173\Desktop\PPT演示"
    copy_excel(folder_path,header=0,sheet_name=[0,1,2,3,4])







