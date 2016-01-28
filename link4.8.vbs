' Скрипт по созданию ярлыков. Версия: 4.8
' Выводит ошибки на экран и в командную строку.
' Создает основные ярлыки
' Создает расшареную папку для обмена и добавляет в нее разрешения
' Копирует url-ярлыки с сервера в избранное
'Работает для Windows 7
'Для ХР надо писать Мои Документы, при создании расшаренной папки

' проверка сервера сценариев. Необходимо для того, что бы работал вывод StdOut в командную строку. Окошки неудобны(
Dim objFS, objWShell, strTranslator
Set objFS = CreateObject("Scripting.FileSystemObject")
strTranslator = objFS.GetBaseName(WScript.FullName)
If StrComp(strTranslator, "wscript", vbTextCompare) = 0 Then
    WScript.Echo "Сервер сценариев по умолчанию: " & UCase(strTranslator) & vbNewLine & "Работаем в режиме перезапуска с другим сервером сценарие."
    Set objWShell = CreateObject("WScript.Shell")
    objWShell.Run "cmd /c echo off && chcp && cscript.exe " & WScript.ScriptFullName, 1 ' cmd /c echo off && 
	Wscript.Quit 0
Else
    WScript.Echo "Сервер сценариев по умолчанию: " & UCase(strTranslator) & vbNewLine & "Работаем в штатном режиме."
End If
Set objWShell = Nothing: Set objFS = Nothing

'set objWShell = CreateObject("Wscript.Shell")
'objWShell.AppActivate "Command Prompt"
'Wscript.Echo Wscript.Version
'objWShell.SendKeys "color 8~ pause~"
'Wscript.Sleep(3000)
'Wscript.StdOut.Write "Test"
'Wscript.Sleep(2000)
'Wscript.Quit 0
'objWShell.SendKeys "exit~"

Dim NoErrors ' флаг ошибок
NoErrors = true

'#####################################################Function Block####################################

'--------------------------#Функция вывода ошибок в поп-ап окне#---------------------------------------------------
Function DispErr(NErr, DErr) 'Вывод ошибок в поп-ап окне
	NoErrors = False
	oShell.Popup "Код: "& NErr & vbNewLine & DErr & vbNewLine, , Wscript.ScriptFullName & ". Error", 0 + 16
End Function
'----------------------------------------------#конец функции#--------------------------------------------

'----------------#Функция добавления разрешений NTFS (вкладка "безопасность")#-----------------------------------------
Function Set_RWEAccess(strDom, strComp, strSAN, strDir)
Dim objWMI, objSecSettings, objSD, objACE
Dim xRes, arrACE, objCollection, objItem, strSID
Dim objSID, objTrustee, objNewACE
Dim blnHasACE, i
Const SE_DACL_PROTECTED = 4096
Const ACCESS_ALLOWED_ACE_TYPE = 0
'Const FULL_ACCESS = 2032127
Const READ_WRITE_EXECUTE_MODIFY = 1245631
Const OBJECT_INHERIT_ACE = 1
Const CONTAINER_INHERIT_ACE = 2
Const INHERITED_ACE = 16

On Error Resume Next
xRes = 0
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComp & "\root\cimv2")
If Err.Number = 0 Then
    Set objSecSettings = objWMI.Get("Win32_LogicalFileSecuritySetting.Path='" & strDir & "'")
    If Err.Number = 0 Then
        If objSecSettings.GetSecurityDescriptor(objSD) = 0 Then
            If Not IsNull(objSD.DACL) Then
                If Not CBool(objSD.ControlFlags And SE_DACL_PROTECTED) Then
                    objSD.ControlFlags = objSD.ControlFlags + SE_DACL_PROTECTED
                    xRes = objSecSettings.SetSecurityDescriptor(objSD)
                End If
                If xRes = 0 Then
                    If Len(strDom) > 0 Then
                        Set objCollection = objWMI.ExecQuery("SELECT SID FROM Win32_Account WHERE Domain='" & strDom & "' AND Name='" & strSAN & "'")
                    Else
                        Set objCollection = objWMI.ExecQuery("SELECT SID FROM Win32_Account WHERE Name='" & strSAN & "'")
                    End If
                    If objCollection.Count > 0 Then
                        For Each objItem In objCollection
                            strSID = UCase(objItem.SID)
                        Next
                        Set objItem = Nothing
                        For Each objACE In objSD.DACL
                            If UCase(objACE.Trustee.SIDString) = strSID Then
                                blnHasACE = True
                                objACE.AceType = ACCESS_ALLOWED_ACE_TYPE
                                objACE.AccessMask = READ_WRITE_EXECUTE_MODIFY
                            End If
                        Next
                        xRes = objSecSettings.SetSecurityDescriptor(objSD)
                        If xRes = 0 Then
                            If Not blnHasACE Then
                                arrACE = objSD.DACL
                                Set objSID = objWMI.Get("Win32_SID.SID='" & strSID & "'")
                                Set objTrustee = objWMI.Get("Win32_Trustee").Spawninstance_()
                                objTrustee.Domain = strDom
                                objTrustee.Name = strSAN
                                objTrustee.SID = objSID.BinaryRepresentation
                                objTrustee.SidLength = objSID.SidLength
                                objTrustee.SIDString = strSID
                                Set objSID = Nothing
                                Set objNewACE = objWMI.Get("Win32_Ace").Spawninstance_()
                                objNewACE.AceType  = ACCESS_ALLOWED_ACE_TYPE
                                objNewACE.AceFlags = OBJECT_INHERIT_ACE + CONTAINER_INHERIT_ACE
                                objNewACE.AccessMask = READ_WRITE_EXECUTE_MODIFY
                                objNewACE.Trustee = objTrustee
                                Set objTrustee = Nothing
                                i = UBound(arrACE) + 1
                                ReDim Preserve arrACE(i)
                                Set arrACE(i) = objNewACE
                                objSD.DACL = arrACE
                                Set objNewACE = Nothing
                                Erase arrACE
                                xRes = objSecSettings.SetSecurityDescriptor(objSD)
                            End If
                        Else
                            xRes = -3
                        End If
                    Else
                        xRes = -2
                    End If
                    Set objCollection = Nothing
                Else
                    xRes = -1
                End If
            Else
                xRes = "Список управления доступом (ACL) к заданному объекту пуст."
            End If
        Else
            xRes = "Не удалось прочитать дескриптор безопасности объекта."
        End If
        Set objSD = Nothing
        Set objSecSettings = Nothing
    Else
        xRes = "Ошибка " & CStr(Err.Number) & vbNewLine & Err.Description
        Err.Clear
    End If
Else
    xRes = "Ошибка " & CStr(Err.Number) & vbNewLine & Err.Description
    Err.Clear
End If
Set objWMI = Nothing
On Error GoTo 0
Set_RWEAccess = xRes
End Function
'----------------------------------------------#конец функции#--------------------------------------------
'#################################################Function Block End########################################

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Main Program~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'создание ярлыков на рабочем столе
Set oShell = WScript.CreateObject("WScript.Shell")
Dim pDesk, sUName
pDesk = oShell.SpecialFolders("Desktop") ' путь к рабочему столу
sUName = oShell.ExpandEnvironmentStrings("%USERNAME%") 'имя пользователя
' Создаем массивы с путями и переменными
Dim n 'количество элементов массива
n = 4 ' массив начинается с 0
Dim aName(), aPath() 'Массивы с именами и путями
ReDim aName(n), aPath(n)
aName(0) = "fserver"
aPath(0) = "\\" & aName(0)
aName(1) = "СЛУЖЕБНАЯ_ПАПКА_МАСКОМ"
aPath(1) = aPath(0) & "\mascom\" & aName(1)
aName(2) = "CПРАВОЧНИК СОТРУДНИКА КОМПАНИИ"
aPath(2) = aPath(1) & "\Документы компании и др . инфо\" & aName(2)
aName(3) = sUName & " (10Гб)"
aPath(3) = aPath(0) & "\" & sUName
aName(4) = "Папка для обмена файлами по сети"
aPath(4) = oShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\" & aName(4)

' создаем папку для обмена файлами
'name="Папка для обмена файлами по сети"
'fPath = oShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\" & name

for i=0 to n '----------------------#цикл для массивов#--------------
if i=4 then '-------#if для обмена файлами#---------------
	Set fso=WScript.CreateObject("Scripting.FileSystemObject") 
	' Если папки не существует, то создаем папку
	if Not fso.FolderExists(aPath(i)) then fso.CreateFolder(aPath(i))
	
	'Расшариваем папку
	dim objSD, objACE
	Const FILE_SHARE = 0
	Const MAXIMUM_CONNECTIONS = 25
	Const ACCESS = 1245631 'маска на чтение и изменение в разрешениях общего доступа
	strComputer = "."
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set objNewShare = objWMI.Get("Win32_Share")
	errReturn = objNewShare.Create (aPath(i), aName(i), FILE_SHARE, MAXIMUM_CONNECTIONS)
	if errReturn=0 then
		set objSecSettings = objWMI.Get("Win32_LogicalShareSecuritySetting.Name='" & aName(i) & "'")
		If objSecSettings.GetSecurityDescriptor(objSD) = 0 Then
				If Not IsNull(objSD.DACL) Then
					For Each objACE In objSD.DACL
						objACE.AccessMask = ACCESS
					Next
					Set objACE = Nothing
					errReturn = objSecSettings.SetSecurityDescriptor(objSD)
					Select Case errReturn
						Case 0: errDescription = "Успешное завершение. Папка расшарена"
						Case 2: errDescription = "Отсутствует доступ к необходимой информации." 
							DispErr errReturn, errDescription
						Case 9: errDescription = "Для выполнения операции недостаточно полномочий."
							DispErr errReturn, errDescription
						Case 21: errDescription = "Заданы недопустимые значения параметров."
							DispErr errReturn, errDescription
						Case Else: errDescription = "Неизвестная ошибка с кодом: " & errReturn
							DispErr errReturn, errDescription
					End Select
				Else
					errDescription = "Список управления доступом к ресурсу " & UCase(aName(4)) & " пуст."
					DispErr errReturn, errDescription
				End If
		Else
			errDescription = "Не удалось прочитать дескриптор безопасности ресурса " & UCase(aName(4))
			DispErr errReturn, errDescription
		End If
		Set objSD = Nothing
		Set objSecSettings = Nothing
	ElseIf errReturn=22 then
		errDescription = "Ошибка " & errReturn & ": Общий ресурс " & UCase(aName(4)) & " уже существует"
		DispErr errReturn, errDescription
	Else
		errDescription = "Ошибка " & errReturn & " при создании ресурса общего доступа " & UCase(aName(4))
		DispErr errReturn, errDescription
	End If
	Wscript.Echo errDescription

	'Добавление пользоватля в разрешения NTFS
	Dim objWsNet, strDomain, strComputer, strAccount
	Dim strPath, xResult, xErr

	strAccount = "Все" 'Пользователь, которого необходимо добавить

	If StrComp(strAccount, "Система", vbTextCompare) = 0 Then strAccount = "System"
        strPath = aPath(i)
        Set objWsNet = CreateObject("WScript.Network")
        strComputer = objWsNet.ComputerName
        Set objWsNet = Nothing        
        If StrComp(strAccount, "System", vbTextCompare) <> 0 And StrComp(strAccount, "Все", vbTextCompare) <> 0 Then
            strDomain = strComputer
        Else
            strDomain = vbNullString
        End If
        xErr = Set_RWEAccess(strDomain, strComputer, strAccount, strPath)
        If IsNumeric(xErr) Then xErr = CStr(xErr)
        Select Case xErr
            Case "-3": xResult = "Не удалось настроить параметры доступа существующей записи " & UCase(strDomain & "\" & strAccount)
            Case "-2": xResult = "Не найдена учётная запись объекта " & UCase(strDomain & "\" & strAccount)
            Case "-1": xResult = "Не удалось отключить наследование безопасности у папки " & UCase(strPath)
            Case "0": Wscript.Echo "Успешное завершение. Права NTFS добавлены"
            Case "2": xResult = "Доступ запрещён."
            Case "8": xResult = "Неизвестная ошибка."
            Case "5", "9": xResult = "Для выполнения операции недостаточно полномочий."
            Case "21": xResult = "Заданы недопустимые значения параметров."
            Case Else: WScript.Echo xErr
        End Select
		if xErr <> 0 then 
			Wscript.Echo xErr & ": " & xResult
			DispErr xErr, xResult
		end If
end if '----------#конец if для обмена файлами#------------------

' Создание ярлыков:
On Error Resume Next
Set oShellLink = oShell.CreateShortcut(pDesk & "\" & aName(i) & ".lnk")
' Целевой путь к файлу для которого создаётся ярлык:
oShellLink.TargetPath = aPath(i)
if Err.Number=0 then 
	oShellLink.Save
	Wscript.Echo "Создан ярлык: " & aName(i)
Else
	Wscript.Echo "Код: "& CStr(Err.Number) & vbNewLine & Err.Description & vbNewLine & "Ошибка при создании ярлыка."
End If

next ' --------------------##конеч цикла массивов##------------------------------

'копирование избранного
set cp = oShell.Exec("cmd /q /k echo off")
cp.StdIn.WriteLine "chcp 1251>nul" ' необходимо для вывода отчета в основное окно в нормальной кодировке
cp.StdIn.WriteLine "xcopy \\fserver\distr\*.url %USERPROFILE%\Favorites /y"
cp.StdIn.WriteLine "exit"
set TextStream = cp.StdOut
str = TextStream.ReadAll
'Do While Not cp.StdOut.AtEndOfStream
   ' strText = cp.StdOut.ReadLine()
	'Wscript.Echo strText
        'Exit Do

'Loop
Wscript.Sleep(2000)
Wscript.Echo str
oShell.Exec "cmd /q /c chcp 866>nul"
WScript.Sleep(1000) 'необходимо для смены кодировки
Wscript.Echo "Скрипт закончил работу"

Wscript.Sleep(5000)