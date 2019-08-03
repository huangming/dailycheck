Public Sub log(message)
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fileSystemObj, fileSpec
    Dim currentTime
    currentDate = Date
    currentTime = Time
    Set fileSystemObj =CreateObject("Scripting.FileSystemObject")
    fileSpec = replace(fileSystemObj.GetFile(Wscript.scriptfullname).path,".vbs", ".log")
    If Not (fileSystemObj.FileExists(filespec)) Then 
        Set logFile = fileSystemObj.CreateTextFile(fileSpec, ForWriting, True) 
        logFile.Close 
        Set logFile = Nothing
    End If
    Set logFile = fileSystemObj.OpenTextFile(fileSpec, ForAppending, False, True)
    logFile.WriteLine ("["&currentDate &" "&currentTime & "] " & message)
    logFile.Close
    Set logFile = Nothing
    Set fileSystemObj = Nothing
End Sub

Function Test(string, Exp)
    '测试字符串中是否包含数字："\d+"
    '测试字符串中是否全是由数字组成："^\d+$"
    '测试字符串中是否有大写字母："[A-Z]+"     
    '测试字符串中是否包含中文："\中文+"
     Set re = New RegExp
    '设置是否匹配大小写
     re.IgnoreCase = False
     re.Pattern = Replace(Exp,"\中文","[\u4E00-\u9FA5]")
     Test = re.Test(string)
     set re = Nothing
End Function

Function FilesTree(sPath,ifchild)  
'遍历一个文件夹下的所有文件夹文件夹  
    dim result
    result = ""
    Set oFso = CreateObject("Scripting.FileSystemObject")  
    If not oFso.folderExists(sPath) Then 
	    FilesTree = ""
        Set oFso = Nothing  
	    Exit Function
    End if
    Set oFolder = oFso.GetFolder(sPath)  
    Set oSubFolders = oFolder.SubFolders  
      
    Set oFiles = oFolder.Files  
    For Each oFile In oFiles  
        result = result + VBcrlf + oFile.Path  
        'oFile.Delete  
    Next  
      
    if ifchild then
        For Each oSubFolder In oSubFolders  
            'result = result + "|" + oSubFolder.Path  
            'oSubFolder.Delete  
            result = result + VBcrlf + FilesTree(oSubFolder.Path, ifchild)'递归  
        Next  
    end if
    FilesTree = right(result, len(result)-2)
    Set oFolder = Nothing  
    Set oSubFolders = Nothing  
    Set oFso = Nothing  
End Function  


Function is_file_exists(sPath, pattern, ifchild)  
    tmp_files = split(FilesTree(sPath,ifchild), VBcrlf)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    dim i:i=0
    for i=0 to Ubound(tmp_files)
        if Test(objFSO.GetFileName(tmp_files(i)), pattern) then
            Set objFSO = Nothing
            is_file_exists = True
            Exit Function
        End If
    next
    is_file_exists = False
End Function

Function formatdate(mydate, pattern)  
    ' 根据日期返回自定义字符串
    if Vartype(mydate) <> 7 then
        mydate = DateValue(mydate)
    end if
    myday = Day(mydate)
    if myday < 10 then myday = "0" + Cstr(myday)
    mymonth = Month(mydate)
    if mymonth < 10 then mymonth = "0" + Cstr(mymonth)
    myyear_all = Year(mydate)
    myyear = right(myyear_all,2)
    formatdate = Replace(pattern, "@yyyy", myyear_all)
    formatdate = Replace(formatdate, "@yy", myyear)
    formatdate = Replace(formatdate, "@mm", mymonth)
    formatdate = Replace(formatdate, "@dd", myday)
End Function
'msgbox formatdate(startdate, "testyy@yyyy@mm@dd")
'msgbox is_file_exists("F:\project\log_zz","data\d{8}.db",True)
sub checkfiles(path, pattern, startdate, enddate, flag)
    'startdate = #22/12/2015#
    'enddate = #01/01/2016#
    'pattern = "test@yyyy@mm@dd.txt"
    'path = "f:\test"
    'flag = "11111100" 第一个1表示是否文件，后面的1表示对应星期数是否备份
    'call checkfiles(path, pattern, startdate, enddate, flag)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")  
    If fso.folderExists(path) Then 
        log("check dir " + path)
    else
        Set fso = Nothing  
        log("dir " + path + " not exist")
	    Exit Sub
    End if
    if startdate = "" then startdate = DateSerial(year(Date), 1, 1)
    if enddate = "" then enddate = Date
    for i = 0 to DateDiff("d", startdate, enddate)
        myday = DateAdd("d", -i, enddate)
        myweek = Weekday(myday, 2)
        if Mid(flag, myweek + 1, 1) = "1" then 
            filename = formatdate(myday, pattern)
            if Mid(flag, 1, 1) = "1" then
                if is_file_exists(path, filename, False) then
                    log("file " + filename + " exist")
                else
                    log("file " + filename + " not found")
                end if
            else
                if fso.folderExists(path + "/" + filename) then
                    log("dir " + filename + " exist")
                else
                    log("dir " + filename + " not found")
                end if
            end if
        end if
    next
    Set fso = Nothing  
end sub

'dim task(3)
dim task(8)
'task(0) = "f:\test|@yyyy@mm@dd|22/12/2015|01/01/2016|01111100"
'task(1) = "f:\test|test@yyyy@mm@dd|22/12/2015|01/01/2016|11111111"
task(0) = "f:\data\jzjy|@yyyy@mm@dd_1.ldm|22/12/2015|01/01/2016|11111100"
task(1) = "f:\data\rzrq|run@yyyy@mm@dd_1.ldm|22/12/2015|01/01/2016|10111100"
task(2) = "f:\caiwu|hl_@yyyy@mm@dd.dmp|22/12/2015|01/01/2016|11111100"
task(3) = "f:\tyzh|after_@yyyy@mm@dd_|22/12/2015|01/01/2016|11111100"
task(4) = "f:\zcgl\hsfa|HSFA_@yyyy@mm@dd.rar|22/12/2015|01/01/2016|11111100"
task(5) = "f:\zcgl\hsorgan|hswinrun@yyyy@mm@dd|22/12/2015|01/01/2016|11111100"
task(6) = "e:\zcgl\hsta|FLast@yy@mm@dd.dmp|22/12/2015|01/01/2016|11111100"
task(7) = "f:\xinyi|FULL_@yyyy@mm@dd.DMP|22/12/2015|01/01/2016|11111100"
task(8) = "e:\zcgl|hswinsql@yyyy@mm@dd.bak|22/12/2015|01/01/2016|11111100"
sub dotask()
    for i = 0 to Ubound(task)
        tmp_arr = split(task(i), "|")
        call checkfiles(tmp_arr(0),tmp_arr(1),DateValue(tmp_arr(2)),DateValue(tmp_arr(3)),tmp_arr(4))
    next
end sub

call dotask()

