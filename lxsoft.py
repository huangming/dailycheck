#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Last Update:

import win32com.client
import pywintypes
import pythoncom
import os

def GoodString(value):
    """
    >>> GoodString(2)
    '2'
    """
    try:
        return str(value)
    except UnicodeEncodeError:
        return value

class Lxsoft():
    def __init__(self, log, path='',visual=0, flag=0):
        self.log = log
        self.lx = self.__registerLxsoft()
        self.InitExcel()
        self.path = path
        if path:
            self.Open(path, visual, flag)#flag=1以只读模式打开
            self.SheetIndex()
            self.SheetCount()

    def __registerLxsoft(self):
        '''注册插件
        '''
        pythoncom.CoInitialize()
        command = "regsvr32 /s atl.dll"
        os.system(command)
        try:
            lx = win32com.client.Dispatch('Lazy.LxjExcel')
        except pywintypes.com_error:
            command = 'regsvr32 /s LazyOffice.dll'
            if os.system(command) == 0:
                lx = win32com.client.Dispatch('Lazy.LxjExcel')
                # assert(dm.Ver == '2.1138')
                return lx
            else:
                self.log.error('regsver32 error')
        else:
            return lx
     
    def InitExcel(self):
        self.path = ''
        self.maxsheet = 1
        self.index = 1
        self.sindex = 1

    def __del__(self):
        try:
            self.Close()
        except:
            pass

    def __Trueindex(self, index):
        if index == '':
            return self.index
        else:
            return index

    def Open(self, path, visual=0, flag=0):
        # visual:1可见,0不可见 opencd:打开密码(如果存在)
        # writecd:写入密码 flag:只读方式打开
        self.path = path
        self.index =  self.lx.ExcelOpen(path, visual, flag)
        return self.index

    def Close(self, index=''):
        index = self.__Trueindex(index)
        return self.lx.ExcelClose(index)

    def SheetIndex(self, index=''):
        # 获取当前标签页序号
        index = self.__Trueindex(index)
        self.sindex = self.lx.SheetIndex(index)
        return self.sindex

    def SheetName(self, Sindex, index=''):
        # 获取标签页序号(或名称)为Sindex的标签名称(或索引)
        index = self.__Trueindex(index)
        return self.lx.SheetGetName(Sindex, index)

    def SheetCount(self, index=''):
        # 获取标签页总数
        index = self.__Trueindex(index)
        self.maxsheet =  self.lx.SheetCount(index)
        return self.maxsheet

    def SheetAdd(self, Sindex, index=''):
        # 在第Sindex个标签页之前新建一个标签
        # 如果Sindex是字符串,则在最后新建一个名叫Sindex的标签
        index = self.__Trueindex(index)
        self.log.debug(u'新建标签: '+GoodString(Sindex))
        return self.lx.SheetAdd(Sindex, index)

    def SheetRename(self, Sindex, name, index=''):
        # Sindex可为序号或者名称
        index = self.__Trueindex(index)
        return self.lx.SheetRename(Sindex, name, index)

    def SheetDel(self, Sindex, index=''):
        index = self.__Trueindex(index)
        self.log.debug(u'删除标签: '+GoodString(Sindex))
        return self.lx.SheetDel(Sindex, index)

    def Write(self, Sindex, x, y, string, index=''):
        # Write(1,3,2,"内容",Index)向单元格(3,2)即'B3'写入内容
        index = self.__Trueindex(index)
        if string is None: string=''
        self.log.debug(u'向单元格(%s,%s)写入内容: %s' % (GoodString(x),GoodString(y),GoodString(string)))
        return self.lx.ExcelWrite(Sindex, x, y, string, index)

    def WriteEx(self, Sindex, ranges, datalist, index=''):
        # WriteEx(1,'A1:B2',[(1,2),(3,4)],Index)
        index = self.__Trueindex(index)
        # self.log.debug(u'向区域 %s 写入内容: %d 条' % (ranges,len(datalist)))
        self.log.debug(u'向区域 %s 写入内容: %d 条' % (ranges,len(datalist)))
        return self.lx.ExcelWriteEx(Sindex, ranges, datalist, index)


    def WriteEx_(self, Sindex, ranges, datalist, index=''):
        # WriteEx(1,'A1:B2',[(1,2),(3,4)],Index)
        # 处理None数据
        newdatalist = []
        for t in datalist:
            if type(t)==type((1,)) or type(t)==type([]):
                newdatalist.append(tuple(['' if x is None else x for x in t]))
            else:
                if t is None:
                    newdatalist.append('')
                else:
                    newdatalist.append(t)
        index = self.__Trueindex(index)
        # self.log.debug(u'向区域 %s 写入内容: %d 条' % (ranges,len(datalist)))
        self.log.debug(u'向区域 %s 写入内容: %d 条' % (ranges,len(newdatalist)))
        return self.lx.ExcelWriteEx(Sindex, ranges, newdatalist, index)

    def CopyTo(self,srcSindex, srcrange, desSindex, desrange, flag=2):
        # 【参数1】来源标签页，整数型的标签索引号或者字符串型旧的标签名称,为0时仅粘贴
        # 【参数2】复制区域，格式为"区域,EXCEL索引号”，区域为""时整表复制
        # 【参数3】去向标签页，整数型的标签索引号或者字符串型旧的标签名称,为0时仅复制
        # 【参数4（可选）】粘贴区域，格式为同复制区域,默认为"A1:A1”，区域为""时整表粘贴
        # 【参数5（可选）】复制方式，默认普通复制,填1相当于格式刷,2为只复制值 
        return self.lx.ExcelCopyTo(srcSindex, srcrange, desSindex, desrange, flag)

    def Save(self, index=''):
        index = self.__Trueindex(index)
        self.lx.ExcelSave(index)

    def Read(self, Sindex, x, y, index=''):
        index = self.__Trueindex(index)
        string = self.lx.ExcelRead(Sindex, x, y, index)[0]
        self.log.debug(u'读取单元格(%s,%s): %s' % (GoodString(x),GoodString(y),GoodString(string)))
        return string

    def ReadEx(self, Sindex, ranges, index=''):
        index = self.__Trueindex(index)
        tmp = self.lx.ExcelReadEX(Sindex, ranges, index)
        if tmp and tmp[0]:
            result = list(tmp[0])
            self.log.debug(u'读取区域 %s 数据: %d 条' % (ranges,len(result)))
            return result
        else:
            self.log.debug(u'读取区域 %s 数据: 0 条' % (ranges))
            return []


    def ReadAll(self, Sindex, startrow=1, index=''):
        index = self.__Trueindex(index)
        result = list(self.lx.ExcelReadEX(Sindex, ranges, index)[0])
        #self.log.debug(u'读取单元格(%s,%s): %s' % (GoodString(x),GoodString(y),GoodString(string)))
        return result

    def differ(self, Sindex, x1, y1, x2, y2, index1='', index2=''):
        index1 = self.__Trueindex(index1)
        index2 = self.__Trueindex(index2)
        return int(self.Read(Sindex, x1, y1, index1)) - int(self.Read(Sindex, x2, y2, index2))

    def LocToAdd(self, row, column):
        # 将(3,2)转化为"B:3"
        result, _ = self.lx.LocToAdd(row, column)
        return result

    def AddToLoc(self, rangestr):
        # 将"B:3"转化为(3,2)
        _, resultx, resulty = self.lx.AddToLoc(rangestr)
        return resultx, resulty

    def Cells(self, tab1, action, tab2, index=''):
        # 如设置tab2为负数，将只进行复制
        # 如设置tab2为0，将在所有标签页后面新建一个被复制的表
        # 将工作表第1个标签页整表复制到第3个标签页:self.lx.ExcelCells(1, u"复制", 3, Index)
        index = self.__Trueindex(index)
        self.log.debug(u'表%s %s -> %s' % (GoodString(tab1),action,GoodString(tab2)))
        return self.lx.ExcelCells(tab1, action, tab2, index)

    def RowsCount(self, Sindex=1, index=''):
        index = self.__Trueindex(index)
        count,_ = self.lx.SheetRowsCount(Sindex, index)
        self.log.debug('标签 %s 行数：%s'%(str(Sindex), str(count)))
        return int(str(count))

    def Rows(self, tab, row1, action, row2=-1, index=''):
        # 如设置row2为负数，将只进行复制
        # 复制第1个标签页第2行到第8行:self.lx.ExcelRows(1, 2, u"复制", 8, Index)
        # 第2行模糊查找1:self.lx.ExcelRows(1, 2, u"模糊查找", 1, Index)
        index = self.__Trueindex(index)
        self.log.debug(u'行%s %s -> %s' % (GoodString(row1),action,GoodString(row2)))
        return self.lx.ExcelRows(tab, row1, action, row2, index)

    def Columns(self, tab, col1, action, col2, index=''):
        # 如设置col2为负数，将只进行复制
        # 复制第1个标签页第3列到第6列:self.lx.ExcelColumns(1, 3, u"复制", 6, Index)
        # 第3列模糊查找1:self.lx.ExcelColumns(1,3, u"模糊查找", 1, Index)
        index = self.__Trueindex(index)
        self.log.debug(u'列%s %s -> %s' % (GoodString(col1),action,GoodString(col2)))
        return self.lx.ExcelColumns(tab, col1, action, col2, index)

    def Range(self, tab, ranges, action, destination, index=''):
        # 如设置destination为负数，将只进行复制
        # 复制第1个标签页"B2:C5"区域:self.lx.ExcelRange(1, "B2:C5", "复制", -1, Index)
        # 区域C2:E5模糊查找1:self.lx.ExcelRange(1, "C2:E5",u"模糊查找", 1, Index)
        # lx.ExcelRange(1,"A1:D5","清除","全部",Index)区域清除,
        # 可设为"格式"、"批注"、"全部"或者"其他"字符串,填其他默认只清除内容
        index = self.__Trueindex(index)
        self.log.debug(u'区域 %s %s -> %s' % (GoodString(tab),action,GoodString(destination)))
        return self.lx.ExcelRange(tab, ranges, action, destination, index)


if __name__ == '__main__':
    import log
    import time
    from pprint import pprint
    # path = 'F:\\project\\tmp\\test.xls'
    # path = 'E:\\OneDrive\\work\\document\\project\\运营资管excel\\20170518华林提取001.xlsx'
    # path = 'E:\\OneDrive\\work\\document\\project\\运营资管excel\\票据模板.xlsx'
    path = 'H:\\OneDrive\\work\\document\\project\\运营资管excel\\20171107浙商总行-华林证券001期.xlsx'
    path = 'H:\\OneDrive\\work\\document\\project\\运营资管excel\\票据模板v2.xlsx'
    #path = 'E:\\onedrive\\python\\python3\\tmp.xlst'
    mylog = log.Log()
    lx = Lxsoft(mylog, path)
    # print int(lx.Read(lx.SheetCount(), 5, 7))==int(time.strftime("%Y%m%d"))
    # a = lx.Read(lx.maxsheet, 5, 7)
    a = lx.SheetCount()
    orderid1=[]
    orderid2=[]
    print(a)
    # a = lx.lx.ExcelReadEX(1,"A1:H8",lx.index)
    # a = lx.ReadEx('追加',"A2:L3")
    # pprint(a)
    # a = lx.Read('全量清单', 2, 8)
    a = lx.Read('托收', 2, 6)
    b = lx.Read('托收', 2, 4)
    # a = lx.ReadEx('全量清单',"A2:L3")
    pprint(a)
    pprint(b)
    import tools
    print(tools.getDateDiff(a,b))
    # a = lx.ReadEx('托收',"A2:L3")
    # pprint(a)
    # a = lx.RowsCount('追加')
    # a = lx.lx.ExcelRows(1,2,"格式","",lx.index)
    # lx.lx.ExcelRows(1,8,"格式",a,lx.index)
    # lx.Rows("追加","3:5","清除","其他")
    # lx.lx.ExcelWriteEx("追加","A4:B5",[(1,2),(3,4)],lx.index)
    #a = lx.Read(1, 1, 1)
    # for i in range(1,649):
    #     orderid1.append(lx.Read(lx.maxsheet, i, 3))
    #     orderid2.append(lx.Read(lx.maxsheet, i, 4))
    # for i in range(649,666):
    #     orderid2.append(lx.Read(lx.maxsheet, i, 4))
    # lx = Lxsoft(config.log)
    # index = lx.Open(path, 0)
    # lx.Close(index)


