#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Last Update: 2015/06/30 10:53:52

#用户文件
tarket_file = 'dailycheck.ini'
#初始化数据, 仅供内部使用
tarket_data = {
    "sys":{"advance_date":"1"},
    "zg":{"host":"127.0.0.1",  
          "user":"root", 
          "pass":"password",
          "backup_dir":"\\test_zg",
          "advance_date":"1",
          "regular":"20%y%m%d"},
     "rzrq":{"host":"127.0.0.1",  
             "user":"root1", 
             "pass":"password",
             "backup_dir":"\\test_rzrq\\",
             "advance_date":"1",
             "regular":"%y%m%d"}
}

import os
def loadData(path=None, init_data='', strcode=""):
    import configparser
    if path == None:
        path = getattr(sys.modules['__main__'], '__file__', 'config.ini')
        path = os.path.basename(path.replace('.py', '.ini'))
    if os.path.isfile(path):
        data = {}
        conf = configparser.ConfigParser()
        conf.read(path, strcode)
        for section in conf.sections():
            # print(conf.items(section))
            db = dict(conf.items(section))
            data.setdefault(section, db)
    else:
        data = init_data
    return data

data = loadData(tarket_file, tarket_data, strcode = "utf-8-sig")


if __name__ == '__main__':
    # log.error('test error')
    # log.debug('test debug')
    # log.info('test info')
    print(data['sys']['excel_path'])
    # for (task, t) in data:
    #     print task+'        '+t
    # for (task, t) in data.items():
        # print task+'        '+t
    print(data)

