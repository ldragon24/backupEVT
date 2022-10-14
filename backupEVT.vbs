Set oSH = Wscript.CreateObject("Wscript.Shell")
Set oFS = CreateObject("Scripting.FileSystemObject")

'День, месяц и год
Dim val1,val2,val3
val1 = Day(Date)
val2 = Month(Date)
val3 = Year(Date)

'Путь для сохранения журналов (путь запуска скрипта)
sPath = oFS.GetParentFolderName(Wscript.ScriptFullName)
sPath =sPath & "\ARCHIVE_EVT"

If Not oFS.FolderExists(sPath) Then
   oFS.CreateFolder(sPath)
End If

'Объект (в данном случае локальный компьютер)
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate, (Backup, Security)}!\\"  & strComputer & "\root\cimv2")
Set colLogFiles = objWMIService.ExecQuery ("SELECT * FROM Win32_NTEventLogFile")

'Сохраняем все журналы какие есть в системе
For Each objLogfile in colLogFiles
 
 strBackupLog = objLogFile.BackupEventLog (sPath & "\" & objLogFile.LogFileName & ".evt")
 '////Очистка журналов
 'objLogFile.ClearEventLog()
 
Next

'Архивируем

Dim Zip, ArcPath
 
' Где создаем архив
ArcPath = sPath & "\" & val1 & "_" & val2 & "_" & val3 & ".zip"
 
Set Zip = New ZipClass
 
 if (Zip.CreateArchive (ArcPath)) then ' старый архив затирается
   ' Zip.CopyFolderToArchive sPath
	Set objShellApp = CreateObject("Shell.Application")
	Set objFolder = objShellApp.NameSpace(sPath & "\")
	Set objFolderItems = objFolder.Items()
	
	'Добавляем все файлы журналов в архив
	objFolderItems.Filter 64+128, "*.evt"
For Each file in objFolderItems
    Zip.CopyFileToArchive file
Next
	
end if
 
 'Удаляем все файлы журналов из папки
strSourceFolder = sPath & "\"

If oFS.FolderExists(strSourceFolder) Then
	oFS.DeleteFile oFS.BuildPath(strSourceFolder, "*.evt"), True
Else
	WScript.Echo "Can't find source folder [" & strSourceFolder & "]."
	WScript.Quit 1
End If

'Удаляем старые архивы
fExt = "zip"              ' Расширение файлов
nMin = 12                 ' Минимальное число оставляемых файлов
nOld = 365                ' Старше кол-ва дней файлы удаляем

OldDate = DateAdd("d", -nOld, Date)

Set Folds = oFS.GetFolder(sPath)
Set Files = Folds.Files

N = Files.Count - 1

If N < 0 Then
    
Else
    ReDim nFiles(N), dFiles(N)
    NN = -1
    For Each jf In Files
        nFiles(NN + 1) = jf.Name
        If LCase(oFS.GetExtensionName(sPath + "\" + nFiles(NN + 1))) = LCase(fExt) Then
            NN = NN + 1
            dFiles(NN) = jf.DateLastModified
        End If
    Next
    
	If NN < 0 Then
       
    Else
        For i = 0 To NN
            For j = i To NN
                If dFiles(i) < dFiles(j) Then
                    df = dFiles(i)
                    dFiles(i) = dFiles(j)
                    dFiles(j) = df
                    nf = nFiles(i)
                    nFiles(i) = nFiles(j)
                    nFiles(j) = nf
                End If
            Next
        Next
        If NN > nMin - 1 Then
            For i = nMin To NN
                If dFiles(i) < OldDate Then Call oFS.DeleteFile(sPath + "\" + nFiles(i), True)
            Next
        End If
    End If
   
End If

Class ZipClass
 
        Private oShApp, oFSO, oArchive, ArcItemsNewCount, oFolderItems, oFolderItem, oArchiveItems, oTarget, oTargetItems, ZipHeader, isEmptyFolder, SHCONTF_FILES_AND_FOLDERS
 
        Private Sub Class_Initialize() 'Инициализация объектов
 
            Const SHCONTF_FOLDERS               = &H20
 
            Const SHCONTF_NONFOLDERS            = &H40
 
            Const SHCONTF_INCLUDEHIDDEN         = &H80
 
            Const SHCONTF_INCLUDESUPERHIDDEN    = &H10000 ' Windows 7 and Later
 
            SHCONTF_FILES_AND_FOLDERS = SHCONTF_FOLDERS Or SHCONTF_NONFOLDERS Or SHCONTF_INCLUDEHIDDEN Or SHCONTF_INCLUDESUPERHIDDEN
 
            Set oShApp = CreateObject("Shell.Application")
 
            set oFSO = CreateObject("Scripting.FileSystemObject")
 
        End Sub
 
        Function UnpackArchive(SourceArchive, DestPath) 'Распаковка архива
 
            Set oArchiveItems = oShApp.NameSpace(SourceArchive).Items
 
            on error resume next
 
            if not oFSO.FolderExists(DestPath) then oFSO.CreateFolder(DestPath)
 
            if Err.Number <> 0 then WScript.Echo("Не хватает прав для создания временной папки распаковки!"): UnpackArchive = false: Exit Function
 
            on error goto 0
 
            Set oTarget = oShApp.NameSpace(DestPath)
 
            set oTargetItems = oTarget.Items
 
            Dim oSCR: set oSCR = CreateObject("Scripting.Dictionary"): oSCR.CompareMode = 1
 
            for each oFolderItem in oTargetItems: oSCR.Add oFolderItem.Name, "": Next ' подсчет кол-ва уникальных файлов
 
            for each oFolderItem in oArchiveItems
 
                if not oSCR.Exists(oFolderItem.Name) then oSCR.Add oFolderItem.Name, ""
 
            Next
 
            oTarget.CopyHere oArchiveItems, 4+16 '(4 - no ProgressBar, 16 - Yes to all, 1024 - suppress all errors)
 
            Do: Wscript.Sleep 200: oTargetItems.Filter SHCONTF_FILES_AND_FOLDERS, "*": Loop Until oTargetItems.Count => oSCR.Count
 
            UnpackArchive = true: set oArchiveItems = Nothing: set oTarget = Nothing
 
        End Function
 
        Function CreateArchive(ZipArchivePath) 'Подготовка ZIP-архива
 
            If lcase(oFSO.GetExtensionName(ZipArchivePath)) <> "zip" Then WScript.Echo("Указано неверное расширение для архива!"): Exit Function
 
            ZipHeader = "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
 
            on error resume next
 
            with oFSO.OpenTextFile(ZipArchivePath, 2, True)
 
                if Err.Number <> 0 then WScript.Echo("Не хватает прав для создания архива!"): CreateArchive = False: Exit function
 
            .Write ZipHeader: .Close: end with
 
            on error goto 0
 
            Do: WScript.Sleep(100): Loop until oFSO.FileExists(ZipArchivePath): WScript.Sleep(200) 'выжидаем время, пока ZIP-архив не будет создан
 
            Set oArchive = oShApp.NameSpace(ZipArchivePath): if Not (oArchive is Nothing) Then CreateArchive = True
 
        End Function
 
        Function CopyFileToArchive(srcFilePath) 'Копируем файл в ZIP-архив
 
            ArcItemsNewCount = oArchive.Items.Count + 1
 
            Dim srcFileName: srcFileName = oFSO.GetBaseName(srcFilePath)
 
            for each oFolderItem in oArchive.Items ' Проверяем, существует ли уже такой файл в архиве
 
                if strcomp(oFolderItem.name, srcFileName) = 0 then ArcItemsNewCount = oArchive.Items.Count - 1: exit for
 
            next
 
            oArchive.CopyHere srcFilePath ', 4 + 16 + 1024 'these options works only with unzipped folder
 
            Do: Wscript.Sleep 200: Loop Until oArchive.Items.Count => ArcItemsNewCount 'Выжидаем пока кол-во объектов в ZIP-архиве станет >= копируемым в него
 
        End Function
 
        Function CopyFolderToArchive(srcFolderPath) 'Копируем содержимое папки в ZIP-архив
 
            Dim sFilter: set oFolderItems = oShApp.NameSpace(srcFolderPath).Items
 
            oFolderItems.Filter SHCONTF_FILES_AND_FOLDERS, "*" 'включаем в архив скрытые файлы
 
            For each oFolderItem in oFolderItems ' поиск пустых папок
 
                isEmptyFolder = false
 
                if oFolderItem.IsFolder then if oFolderItem.GetFolder.Items.Count = 0 then isEmptyFolder = true
 
                if not isEmptyFolder then sFilter = sFilter & ";" & replace(oFolderItem.Name, ";", "?") ' белый список объектов для фильтра
 
            Next
 
            oFolderItems.Filter SHCONTF_FILES_AND_FOLDERS, mid(sFilter, 1)
 
            ArcItemsNewCount = oArchive.Items.Count + oFolderItems.Count
 
            oArchive.CopyHere oFolderItems
 
            Do: Wscript.Sleep 200: Loop Until oArchive.Items.Count => ArcItemsNewCount 'Выжидаем пока кол-во объектов в ZIP-архиве станет >= копируемым в него
 
        End Function
 
End Class
