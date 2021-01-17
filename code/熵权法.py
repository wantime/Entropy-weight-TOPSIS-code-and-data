import numpy as np
import xlrd
import xlwt
import string

# 读数据并求熵
path = r'E:\arcpy\经济原始数据.xls'
hn, nc = 1, 1
# hn为表头行数,nc为表头列数
sheetname = u'data'


def readexcel(hn, nc):
    data = xlrd.open_workbook(path)
    table = data.sheet_by_name(sheetname)
    nrows = table.nrows
    ncols = table.ncols
    data = []
    for i in range(nc, ncols):
        data.append(table.col_values(i)[hn:])
    return np.array(data)


def entropy(data0):
    # 返回每个样本的指数
    # 样本数，指标个数
    n, m = np.shape(data0)
    # 一列一个样本，一行一个指标
    # 下面是归一化
    maxium = np.max(data0, axis=0)
    minium = np.min(data0, axis=0)
    data = (data0 - minium) * 1.0 / (maxium - minium)
    ##计算第j项指标，第i个样本占该指标的比重
    sumzb = np.sum(data, axis=0)
    data = data / sumzb
    # 对ln0处理
    a = data * 1.0
    a[np.where(data == 0)] = 0.0001
    #    #计算每个指标的熵

    e = (-1.0 / np.log(n)) * np.sum(data * np.log(a), axis=0)
    #    #计算权重
    x1=1-e
    w = (1 - e) / np.sum(1 - e)
    print (w)
    return w


data = readexcel(hn, nc)
wet = entropy(data)


#outxlspath=arcpy.GetParameterAsText(1)
all_str = string.ascii_letters + string.digits
outxlspath = (r'E:\arcpy\weight_jj.xls')  # 新建excel文件
workbook = xlwt.Workbook(encoding='utf-8')  # 写入excel文件
sheet = workbook.add_sheet('weight', cell_overwrite_ok=True)  # 新增一个sheet工作表
headlist = [u'权重']  # 写入数据头
row = 0
col = 0
for head in headlist:
    sheet.write(row, col, head)
    col = col + 1
for i in range(0, len(wet)):  # 写入14行数据
    _weight = wet[i]
    sheet.write(i+1, 0, _weight)
    workbook.save(outxlspath)  # 保存