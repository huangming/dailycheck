#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Last Update:

'''docstring
'''

import time
import datetime
import os
import win32file
import re
import lxsoft
import config
import log
import tools
import random
import base64

class Task():
    def __init__(self, log, lx, sys, task, data):
        self.log = log
        self.lx = lx
        self.sys = sys
        self.task = task
        self.first_txt_line = 0
        self.max_txt_line = int(self.sys.get("max_txt_line",0))
        self.host = data.get("host")
        self.user = data.get("user","administrator")
        if data.get("ifencrypt",0)=='1' or self.sys.get("ifencrypt",0)=='1':
            # base64.b64encode('password'.encode())[::-1]
            try:
                self.password = base64.b64decode(data.get("pass")[::-1]).decode()
            except:
                self.log.trace()
                self.password = "<class 'binascii.Error'>: Incorrect padding"
        else:
            self.password = data.get("pass")
        # self.log.debug('%s:password=%s'%(self.task,self.password))
        self.isexcel = self.sys.get("excel_path",'') != ''
        self.istxt = self.sys.get("txtlog_path",'') != ''
        self.minspace = int(self.sys.get("minspace",'50'))
        self.tradedateonly = data.get("tradedateonly",'1')
        self.path = data.get("backup_dir")
        self.isfile = data.get("isfile","1")
        self.isSpecial = data.get("isSpecial","0")
        self.pattern = data.get("regular","(\d{8})")
        self.advance_day = data.get("advance_day","1")
        self.today = self.sys.get("sysdate", time.strftime("%Y%m%d"))
        self.number = int(data.get("number"))
        self.import_from = data.get("import",'')
        self.txt_result = []
        self.init_txt_result()
        self.first_task_num = self.sys.get("first_task_num",5)

    def init_txt_result(self):
        if self.istxt:
            f = open(self.sys.get("txtlog_path"),'r')
            self.txt_result = f.readlines()
            f.close()
            self.if_logtxt_today()
            for i in range(100):
                if 'G' in self.txt_result[i] or '20' in self.txt_result[i]:
                    self.first_txt_line = i
                    break

    def if_logtxt_today(self):
        tmp_list = self.txt_result[:5]
        for i in range(100):
            if self.max_txt_line > 0:
                tmp_list = self.txt_result[:self.max_txt_line]
                break
            if len(self.txt_result[i])<4: 
                tmp_list = self.txt_result[:i+2]
                self.max_txt_line = i+3
                break
        for i in range(len(tmp_list),1,-1):
            if self.today in self.txt_result[i]:
                self.log.debug('%s:今天已检查'%self.task)
                self.log.debug('%s:%s %s'%(self.task,self.today,self.txt_result[4]))
                break
        else:
            self.log.debug('%s:今天未检查'%self.task)
            self.txt_result=tmp_list + self.txt_result
            # self.log.debug(self.txt_result)
            f = open(self.sys.get("txtlog_path"),'w')
            f.writelines(self.txt_result)
            f.close()
            # self.log.debug(self.txt_result)

    def Connect(self):
        net_addr = self.host[:-2]+'ipc$'
        if self.import_from != '':
            self.log.debug(self.import_from)
            if os.path.exists(self.import_from):
                self.log.debug("import_from 读取正常")
                return 0
        if os.path.exists(self.host):
            return 0
        if 'localhost' in self.host or '127.0.0.1' in self.host:
            return 0
            # cmd_connect = 'Subst '+self.disk+self.host[-3:-2]+':/'+' >nul 2>nul'
            # cmd_disconnect = 'Subst '+self.disk+'/d'
        cmd_connect = 'net use %s %s /USER:%s >nul 2>nul'%(net_addr,self.password,self.user)
        cmd_disconnect = 'net use '+net_addr+'/delete >nul 2>nul'
        tools.command_run(cmd_disconnect, int(self.sys["timeout"]))
        result = tools.command_run(cmd_connect, int(self.sys.get("timeout","5")))
        # self.log.debug('%s:connect=%s'%(self.task,cmd_connect))
        if result != 0:
            self.log.error(self.task.ljust(8)+u'[错误]连接服务器失败!')
            self.log.debug('%s:connect=%s'%(self.task,str(result)))
        if self.import_from != '':
            if not os.path.exists(self.import_from):
                self.log.error(self.task.ljust(8)+u'[错误]*打开结果文件 '+self.import_from+u' 失败!')
                result = -1
        return result

    def __del__(self):
        pass
        # self.Disconnect()

    def Disconnect(self):
        net_addr = self.host[:-2]+'ipc$'
        cmd_disconnect = 'net use %s /delete >nul 2>nul'%net_addr
        # cmd_disconnect = 'net use '+self.disk+'/delete >nul 2>nul'
        if 'localhost' in self.host or '127.0.0.1' in self.host:
            return 0
            # cmd_disconnect = 'Subst '+self.disk+'/d'
        result = tools.command_run(cmd_disconnect, int(self.sys.get("timeout","5")))
        self.log.debug("%s:disconnect=%s"%(self.task,str(result)))
        return result

    def TruePath(self, path, advance_day):
        truepath = path.replace('%y', self.LackDay(advance_day)[2:4])
        truepath = truepath.replace('%Y', self.LackDay(advance_day)[0:4])
        truepath = truepath.replace('%m', self.LackDay(advance_day)[4:6])
        truepath = truepath.replace('%d', self.LackDay(advance_day)[6:8])
        self.log.debug(self.task+':truepath='+truepath)
        if '.*' in truepath:
            try:
                left_path,right_path = truepath.split('.*\\')
                FileNames=os.listdir(left_path)
                paths = []
                for fn in FileNames:
                    tmp_path = os.path.join(left_path,fn)
                    if os.path.isdir(tmp_path):
                        paths.append(os.path.join(left_path,fn))
                truepath = os.path.join(random.choice(paths),right_path)
            except Exception as err:
                pass
                # self.log.error(u'        [错误]路径'+left_path+u'不存在')
                # self.log.exception(err)
        return truepath

    def LackDateAll(self):
        if self.import_from != '' or self.advance_day[0] != '1':
            return True
        flag = "||" if "||" in self.path else "&&"
        root_path = self.host if self.import_from == '' else (self.host[-2:-1]+':')
        path = os.path.join(root_path, self.path.split(flag)[0])
        pattern = self.pattern.split(flag)[0]
        isfile = self.isfile.split(flag)[0]
        FileNames=os.listdir(path)
        if FileNames == []:
            self.log.error(self.task.ljust(8)+u'[错误]'+path+u'文件为空')
            return 0
        days = []
        for fn in FileNames:
            self.log.debug(pattern)
            day = re.search(pattern, fn, re.IGNORECASE)
            if day != None:
                if not isfile:
                    if not os.listdir(os.path.join(path,day.group(0))):
                       continue 
                tmpday=int(day.group(1).replace('-','').replace('_','').replace('/',''))
                tmpday = tmpday + 20000000 if tmpday < 20000000 else tmpday
                days.append(tmpday)
        if days == []:
            # 对多任务任意成功即可的任务一条任务失败不报错处理
            if '||' not in self.path:
                self.log.error(self.task.ljust(8)+u'[错误]正则错误/无匹配: '+pattern)
            else:
                self.log.debug(self.task.ljust(8)+u'[错误]正则错误/无匹配: '+pattern)
            return 0
        self.log.info(self.task.ljust(8)+u'days:'+repr(days))
        if os.path.exists('tradedate.txt'):
            f = open('tradedate.txt','r')
            tradedate_list=f.readlines()
            f.close()
            lackday = []
            for day in tradedate_list:
                if int(day.replace('\n','')) >= int(self.sys.get('beginday','20170101')) and \
                   int(day.replace('\n','')) <= int(self.sys.get('endday',self.today)):
                       if int(day.replace('\n','')) not in days:
                           self.log.info(self.task.ljust(14)+u'备份不存在: '+day)
                           lackday.append(day)
            if lackday:
                f = open('%s.txt'%self.task,'w')
                f.writelines([str(day) for day in lackday])
                f.close()
            else:
                self.log.info(self.task.ljust(14)+u'备份均存在')

    def LastDate(self, path, pattern, isfile):
        FileNames=os.listdir(path)
        if FileNames == []:
            self.log.error(self.task.ljust(8)+u'[错误]'+path+u'文件为空')
            return 0
        days = []
        for fn in FileNames:
            day = re.search(pattern, fn, re.IGNORECASE)
            if day != None:
                if not isfile:
                    if not os.listdir(os.path.join(path,day.group(0))):
                       continue 
                days.append(int(day.group(1).replace('-','').replace('_','').replace('/','')))
        if days == []:
            # 对多任务任意成功即可的任务一条任务失败不报错处理
            if '||' not in self.path:
                self.log.error(self.task.ljust(8)+u'[错误]正则错误/无匹配: '+pattern)
            else:
                self.log.debug(self.task.ljust(8)+u'[错误]正则错误/无匹配: '+pattern)
            return 0
        last_bk = max(days)
        if last_bk < 999999 and last_bk > 99999:
            last_bk = last_bk + 20000000
        return last_bk

    def DateDiff(self, date1, date2):
        if int(date1) <10000000:
            date1 = int(date1) + 20000000
        if int(date2) <10000000:
            date2 = int(date2) + 20000000
        date1_t = time.strptime(str(date1),'%Y%m%d')
        date2_t = time.strptime(str(date2),'%Y%m%d')
        date1_d = datetime.datetime(*date1_t[:3])
        date2_d = datetime.datetime(*date2_t[:3])
        if not os.path.exists('tradedate.txt') or self.tradedateonly !='1':
            self.log.debug(self.task+':datediff='+str((date2_d - date1_d).days))
            return (date2_d - date1_d).days
        else:
            # 计算两日期间隔的交易日数
            f = open('tradedate.txt','r')
            tradedate_list=f.readlines()
            f.close()
            for i in range(len(tradedate_list)):
                if int(tradedate_list[i])>int(date1):
                    date1_index=i-1
                    break
            for i in range(len(tradedate_list)):
                if int(tradedate_list[i])>=int(date2):
                    date2_index=i
                    break
            self.log.debug(self.task+':datediff='+str(date2_index-date1_index))
            return date2_index-date1_index

    def LackDay(self, advance_day, last_bk=''):
        self.log.debug(self.task+':advance_day='+str(advance_day))
        last_bk_dt = datetime.datetime.strptime(str(last_bk), "%Y%m%d") if last_bk else None
        if not os.path.exists('tradedate.txt') or self.tradedateonly !='1':
            if last_bk_dt and int(advance_day)>2: # Truepath函数使用，目前对advance_day>1并且路径有日期标识符的情况不能处理.
                # 对检查日期小于2的用今天往前推advance_day的方式算缺少备份日期
                # 大于等于2的用最后备份日期往后推advance_day方式计算
                lack_day = last_bk_dt + datetime.timedelta(days = int(advance_day))
            else:
                # lack_day = datetime.datetime.now() - datetime.timedelta(days = int(advance_day))
                lack_day = datetime.datetime.strptime(self.today, "%Y%m%d") - datetime.timedelta(days = int(advance_day))
            return lack_day.strftime("%Y%m%d")
        else:
            # 计算两日期间隔的交易日数
            f = open('tradedate.txt','r')
            tradedate_list=f.readlines()
            f.close()
            for i in range(len(tradedate_list)):
                if last_bk and int(advance_day)>2:
                    if int(tradedate_list[i])>int(last_bk):
                        return tradedate_list[i+int(advance_day)-1].replace('\n','')
                        break
                elif int(tradedate_list[i])>=int(self.today):
                    return tradedate_list[i-int(advance_day)].replace('\n','')
                    break
            else:
                self.log.error('LackDay tradedate_list index error:advance_day=%s;last_bk=%s'%(advance_day, last_bk))
                return 'error'

    def import_result(self):
        f = open(self.import_from,'r')
        txt=f.read()
        # self.log.debug('\n%s'%txt)
        pattern = '\|.*?\|\s*'+self.today+'\s*\|\s*.*?\s*\|\s*'+self.task+'\s*\|'
        self.log.debug(pattern)
        # self.log.debug(txt)
        value = re.findall(pattern,txt)
        if value:
            pattern = r'\|(\s*.*?\s*)\|\s*(.*?)\s*\|\s*(.*?)\s*\|\s*(.*?)\s*\|\s*(.*?)\s*\|(\s*.*?\s*)\|\s*(.*?)\s*\|(\s*.*?\s*)\|(\s*.*?\s*)\|'
            value = re.findall(pattern,value[0])[0]
            if value[1] == u'是':
                self.log.info(self.task.ljust(14)+u'*备份正常,最后备份日期: '+value[2])
            else:
                self.log.warning(self.task.ljust(8)+u'[警告]*备份不存在!缺少备份日期: '+value[3])
            return value[1]==u'是', value[2], value[3], value[4]
        else:
            self.log.error(self.task.ljust(8)+u'[错误]*导入失败,结果中无今日记录!')
            return None, None, None, None

    def Checkfile(self):
        if self.import_from != '':
            return self.import_result()
        # 20170216,新增逻辑与非判断，&&表示所有备份检查都要有才算有备份
        # ||表示只要有一个有备份，就算有备份。
        # 因为下面的逻辑是不管有没有多条件，都用多条件分割，所以配置没有多条件的话，实际
        # 逻辑是有默认多条件的，所以对配置的多条件判断要全部统一。这里全部用||判断，没有||默认为&&
        result = False if "||" in self.path else True
        flag = "||" if "||" in self.path else "&&"
        path_list = self.path.split(flag)
        pattern_list = self.pattern.split(flag)
        isfile_list = self.isfile.split(flag)
        advance_day_list = self.advance_day.split(flag)
        last_bk_max = 0
        last_failbk_max = 0
        lack_day_max = u"无"
        for i in range(len(pattern_list)):
            isfile = int(isfile_list[i]) != 0 and isfile_list[i] != ''
            advance_day = int(advance_day_list[i])
            # path = os.path.join(self.sys["disk"], path_list[i])
            root_path = self.host if self.import_from == '' else (self.host[-2:-1]+':')
            path = os.path.join(root_path, path_list[i])
            path = self.TruePath(path, advance_day)
            pattern = pattern_list[i]
            if not os.path.exists('tradedate.txt') or self.tradedateonly !='1':
                tmp_today = time.strptime(self.today, "%Y%m%d")
                if int(time.strftime('%w', tmp_today)) <= advance_day and advance_day < 7:
                    if int(time.strftime('%w'), tmp_today) == 0:
                        # 星期天
                        advance_day = advance_day + 1
                    elif int(time.strftime('%w'), tmp_today) != 6:
                        advance_day = advance_day + 2

            if not os.path.exists(path):
                self.log.error(self.task.ljust(8)+u'[错误]路径 '+path+u' 不存在!')
                result = False if "||" not in self.path else result
                continue
            space = self.FreeSpace(path)

            last_bk = self.LastDate(path, pattern, isfile)
            if last_bk == 0:
                result = False if "||" not in self.path else result
                continue
            if int(advance_day)>2 and last_failbk_max == 0:
                last_failbk_max = last_bk # 对检查频率大于2的多任务情况，以第一个任务的结果算失败的last_bk
            if last_bk > last_bk_max:
                last_bk_max = last_bk
            if self.DateDiff(last_bk, int(self.today)) > int(advance_day):
                lack_day = self.LackDay(advance_day, last_bk)
                self.log.debug('%s:lackday=%s,path=%s,pattern=%s'%(self.task,lack_day,path,pattern))
                result = False if "||" not in self.path else result
                if int(lack_day) > int(0 if lack_day_max==u"无" else lack_day_max): lack_day_max = lack_day
            else:
                result = True if "||" in self.path else result
        if result:
            self.log.info(self.task.ljust(14)+u'备份正常,最后备份日期: '+str(last_bk_max))
            lack_day = u"无"
            lack_day_max = u"无"
        else:
            # 对最终未备份并且检查频率大于2的多任务情况，以第一个任务的结果算失败的last_bk
            if int(advance_day) > 2:last_bk_max = last_failbk_max
            self.log.warning(self.task.ljust(8)+u'[警告]备份不存在!缺少备份日期: '+str(lack_day_max))
        self.log.debug('%s: result:%s,last_bk:%s,lack_day:%s, space:%s' % (self.task,result, last_bk_max, lack_day_max, space))
        return result, last_bk_max, lack_day_max, space

    def Logresult_excel(self, ifbackup, last_bk, lack_day, space):
        if self.number > 0:
            if ifbackup:
                self.lx.Write(self.lx.SheetCount(),self.number,int(self.sys["if_bk"]),u"是")
                color = self.sys.get('normalcolor',"000000")
                self.lx.Rows(self.lx.SheetCount(),self.number,"字体颜色",color)
            else:
                self.lx.Write(self.lx.SheetCount(),self.number,int(self.sys["if_bk"]),u"否")
                color = self.sys.get('errorcolor',"0000FF")
                self.lx.Rows(self.lx.SheetCount(),self.number,"字体颜色",color)
            self.lx.Write(self.lx.SheetCount(),self.number,int(self.sys["lack_date"]),lack_day)
            self.lx.Write(self.lx.SheetCount(),self.number,int(self.sys["last_bk"]),last_bk)
            self.lx.Write(self.lx.SheetCount(),self.number,int(self.sys["space"]),space)
            self.lx.Write(self.lx.SheetCount(),self.number,int(self.sys["check_y"]),self.today)
            if float(space[:-2])<self.minspace:
                # 磁盘不足单元格字体颜色设置
                cell = self.lx.LocToAdd(self.number, int(self.sys["space"]))
                color = self.sys.get('errorcolor',"0000FF")
                self.lx.Range(self.lx.SheetCount(),cell+':'+cell,"字体颜色",color)


    def Logresult_txt(self, ifbackup, last_bk, lack_day, space):
        if self.number <=0: return False
        self.log.debug('max_txt_line:%s'%self.max_txt_line)
        self.log.debug('first_txt_line:%s'%self.first_txt_line)
        self.log.debug('first_task_num:%s'%self.first_task_num)
        self.log.debug('self.number:%s'%self.number)
        backup_result = u'是' if ifbackup else u'否'
        txt = self.txt_result[self.number*2-self.first_task_num*2+self.first_txt_line]
        pattern = r'\|(\s*.*?\s*)\|\s*(.*?)\s*\|\s*(.*?)\s*\|\s*(.*?)\s*\|\s*(.*?)\s*\|(\s*.*?\s*)\|\s*(.*?)\s*\|(\s*.*?\s*)\|(\s*.*?\s*)\|'
        p = re.compile(pattern)
        self.log.debug(txt)
        old_value = re.findall(pattern,txt)[0]
        self.log.debug(old_value)
        lack_day = lack_day.center(9) if lack_day==u'无' else lack_day.center(10)
        new_value = (old_value[0],backup_result.center(9),str(last_bk).center(10),lack_day,space.center(10),old_value[5],self.today.center(10),old_value[7],old_value[8])
        txt = '|'
        for i in range(9):
            txt = txt+new_value[i]+'|'
        f = open(self.sys.get("txtlog_path"),'w')
        f.writelines(self.txt_result[:self.number*2-self.first_task_num*2+self.first_txt_line]+[txt+'\n']+self.txt_result[self.number*2-self.first_task_num*2+self.first_txt_line+1:])
        f.close()

    def Work(self):
        ifbackup, last_bk, lack_day, space = self.Checkfile()
        if ifbackup is None:
            if self.isexcel:
                color = self.sys.get('warningcolor',"FF0000")
                self.lx.Rows(self.lx.SheetCount(),self.number,"字体颜色",color)
            return False
        self.log.debug(self.task+':lask_bk='+str(last_bk))
        if self.isexcel:
            self.Logresult_excel(ifbackup, last_bk, lack_day, space)
        if self.istxt:
            self.Logresult_txt(ifbackup, last_bk, lack_day, space)
        return True

    def FreeSpace(self,path):
        space = 0
        drv,left = os.path.splitdrive(path)
        sectorsPerCluster, bytesPerSector, numFreeClusters, totalNumClusters \
                   = win32file.GetDiskFreeSpace(drv)
        space = (numFreeClusters*sectorsPerCluster*bytesPerSector)/(1024*1024*1024)
        if space < self.minspace:
            self.log.warning(self.task.ljust(8)+u'[警告]磁盘空间不足%sG '%self.minspace+\
                    self.host[-3]+u'盘:'+'%.1fG'%space)
        return '%.1fGB'%space

class DailyCheck():
    def __init__(self, log, lx, data):
        self.log = log
        self.lx = lx
        self.data = data
        self.sys = data['sys']
        self.log.set_logger(cmdlevel=int(self.sys['log_level']),filelevel=int(self.sys['log_level']))
        if self.sys.get('excel_path','') != '':
            self.lx.Open(self.sys['excel_path'])
        self.today = self.sys.get("sysdate", time.strftime("%Y%m%d"))
        # self.log.debug(data)

    def CheckSheet(self, col):
        # 检测最后一个标签的检查日期是否今天。
        if self.sys.get('excel_path','') != '':
            if int(self.lx.Columns(self.lx.SheetCount(), col, u"模糊查找", self.today)[0][0]) ==0:
                self.lx.Cells(self.lx.SheetCount(), u"复制", 0)

    def Work(self):
        self.CheckSheet(int(self.sys["check_y"]))
        first_task_num = min([int(t["number"]) for (task,t) in self.data.items() if task !='sys'])
        self.sys["first_task_num"] = first_task_num
        for (task, t) in self.data.items():
            if task == 'sys':
                continue
            tarket = Task(self.log, self.lx, self.sys, task, t)
            if tarket.Connect() == 0:
                try:
                    tarket.Work()
                    if self.sys.get('checkall','N').lower()[0] == 'y':
                        tarket.LackDateAll()
                except:
                    self.log.trace()
                    if t.isexcel:
                        color = self.sys.get('warningcolor',"FF0000")
                        self.lx.Rows(self.lx.SheetCount(),tarket.number,"字体颜色",color)
                if t.get("disconnect",'0')=='1':
                    tarket.Disconnect()
        self.log.debug("do_after: %s"%self.sys.get("do_after",''))
        if self.sys.get('do_after','') != '':
            result = tools.command_run(self.sys.get("do_after"), int(self.sys.get("timeout","3")))
            log.debug("do_after: %s"%result)

if __name__ == '__main__':
    print(' ')
    print(50*'-')
    print(' ')
    print(time.strftime("%Y-%m-%d %A",time.localtime()).center(50))
    print(' ')
    print(50*'-')
    print(' ')
    # log = log.Log()
    log = log.Log(cmdlevel = 'info')
    if config.data['sys'].get('excel_path','') != '':
        lx = lxsoft.Lxsoft(log)
    else:
        lx = ''
    w = DailyCheck(log, lx, config.data)
    w.Work()
    if config.data['sys'].get('excel_path','') != '':
        lx.Close()
    if config.data['sys']['autorun'] == '':
        os.system("pause") 
    # try:
    #     w.Work()
    #     lx.Close()
    # except:
    #     pass
    # finally:
    #     lx.Close()
    #     os.system("pause")

