import numpy as np
import xlrd
import xlwt
import string

#path=arcpy.GetParameterAsText(0)
path = r'E:\arcpy\生态原始数据.xls'
hn, nc = 1, 1
sheetname = r'data'


def readexcel(hn, nc):
    data = xlrd.open_workbook(path)
    table = data.sheet_by_name(sheetname)
    nrows = table.nrows
    ncols = table.ncols
    data = []
    for i in range(nc, ncols):
        data.append(table.col_values(i)[hn:])
    return np.array(data)

def read_weight_excel(hn, nc):
    _weight_path = r'E:\arcpy\wetst.xls'
    _weight = xlrd.open_workbook(_weight_path)
    _sheetname=r'weight'
    table = _weight.sheet_by_name(_sheetname)
    nrows = table.nrows
    ncols = table.ncols
    weight = []
    for i in range(nc, nrows):
        weight.append(table.row_values(i)[:hn])
    return np.array(weight).transpose(1,0)

def BZH(data0):
    maxium = np.max(data0, axis=0)
    minium = np.min(data0, axis=0)
    data = (data0 - minium) * 1.0 / (maxium - minium)
    return data

def GFH(_data,_weight):
    data_GFH=_data*_weight
    return data_GFH

def TOPSIS(_data_gfh):
    _Vmax = np.max(_data_gfh, axis=0)
    _Vmin = np.min(_data_gfh, axis=0)
    rows,clos=_data_gfh.shape
    vmax=_data_gfh-_Vmax
    vmin=_data_gfh-_Vmin
    vmaxvmax=vmax*vmax
    vminvmin=vmin*vmin
    sum_vmaxvmax=_Vmax = np.sum(vmaxvmax, axis=1)
    sum_vminvmin=_Vmax = np.sum(vminvmin, axis=1)
    Dmax=np.sqrt(sum_vmaxvmax)
    Dmin=np.sqrt(sum_vminvmin)
    T=Dmin/(Dmax+Dmin)
    return T

data = readexcel(hn, nc)
data_bzh = BZH(data)
weight = read_weight_excel(hn, nc)
data_gfh = GFH(data_bzh,weight)
data_topsis = TOPSIS(data_gfh)
d_topsis=data_topsis.reshape(13,5)

#outxlspath=arcpy.GetParameterAsText(1)
all_str = string.ascii_letters + string.digits
outxlspath = (r'E:\arcpy\topsisST.xls')  # 新建excel文件
workbook = xlwt.Workbook(encoding='utf-8')  # 写入excel文件
sheet = workbook.add_sheet('weight', cell_overwrite_ok=True)  # 新增一个sheet工作表
headlist = [u'贴近度']  # 写入数据头
row = 0
col = 0
for head in headlist:
    sheet.write(row, col, head)
    col = col + 1
d_rows,d_cols=d_topsis.shape
for i in range(0, d_rows):
    for j in range(0, d_cols):
        _top = d_topsis[i,j]
        sheet.write(i+1, j, _top)
    workbook.save(outxlspath)  # 保存
