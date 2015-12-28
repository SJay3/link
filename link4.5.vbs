' ������ �� �������� �������. ������: 4.5
' ������� ������ �� ����� � � ��������� ������.
'�������� ��� Windows 7
'��� �� ���� ������ ��� ���������, ��� �������� ����������� �����

' �������� ������� ���������. ���������� ��� ����, ��� �� ������� ����� StdOut � ��������� ������. ������ ��������(
Dim objFS, objWShell, strTranslator
Set objFS = CreateObject("Scripting.FileSystemObject")
strTranslator = objFS.GetBaseName(WScript.FullName)
If StrComp(strTranslator, "wscript", vbTextCompare) = 0 Then
    WScript.Echo "������ ��������� �� ���������: " & UCase(strTranslator) & vbNewLine & "�������� � ������ ����������� � ������ �������� ��������."
    Set objWShell = CreateObject("WScript.Shell")
    objWShell.Run "cmd /c echo off && cscript.exe " & WScript.ScriptFullName, 1 ' cmd /c echo off && 
	Wscript.Quit 0
Else
    WScript.Echo "������ ��������� �� ���������: " & UCase(strTranslator) & vbNewLine & "�������� � ������� ������."
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

Dim NoErrors ' ���� ������
NoErrors = true



Function DispErr(NErr, DErr)
	NoErrors = False
	oShell.Popup "���: "& NErr & vbNewLine & DErr & vbNewLine, , Wscript.ScriptFullName & ". Error", 0 + 16
End Function

'�������� ������� �� ������� �����
Dim sServer 
sServer = "fserver" ' ���������� ��� ������� � ����������

Set oShell = WScript.CreateObject("WScript.Shell")
' �������� ������ �� fserver:
On Error Resume Next
Set oShellLink = oShell.CreateShortcut(oShell.SpecialFolders("Desktop") & "\" & sServer &".lnk")
' ������� ���� � ����� ��� �������� �������� �����:
oShellLink.TargetPath = "\\" & sServer
if Err.Number=0 then 
	oShellLink.Save
	Wscript.Echo "����� �� �������"
Else
	Wscript.Echo "���: "& CStr(Err.Number) & vbNewLine & Err.Description & vbNewLine & "������ ��� �������� ������."
End If

' ����� �� ��������� ����� ������
On Error Resume Next
set oShellLink = oShell.CreateShortcut(oShell.SpecialFolders("Desktop") & "\���������_�����_������.lnk")
oShellLink.TargetPath = "\\" & sServer & "\mascom\���������_�����_������"
if Err.Number=0 then
	oShellLink.Save
	Wscript.Echo "����� �� ��������� �����"
Else
	Wscript.Echo "���: "& CStr(Err.Number) & vbNewLine & Err.Description & vbNewLine & "������ ��� �������� ������."
	DispErr Err.Number, Err.Description
End If

' ����� �� ���������� ����������
On Error Resume Next
set oShellLink = oShell.CreateShortcut(oShell.SpecialFolders("Desktop") & "\C��������� ���������� ��������.lnk")
oShellLink.TargetPath = "\\" & sServer & "\mascom\���������_�����_������\��������� �������� � �� . ����\C��������� ���������� ��������"
if Err.Number=0 then
	oShellLink.Save
	Wscript.Echo "����� �� ����������"
Else
	Wscript.Echo "���: "& CStr(Err.Number) & vbNewLine & Err.Description & vbNewLine & "������ ��� �������� ������."
	DispErr Err.Number, Err.Description
End If

' ����� �� ���� ����� �� ������� � ������������ � 10 ��
On Error Resume Next
set oShellLink = oShell.CreateShortcut(oShell.SpecialFolders("Desktop") + "\" + oShell.ExpandEnvironmentStrings("%USERNAME%") + " (10 ��).lnk")
oShellLink.TargetPath = "\\" & sServer & "\" & oShell.ExpandEnvironmentStrings("%USERNAME%")
If Err.Number=0 then
	oShellLink.Save
	Wscript.Echo "����� �� ������� �����"
Else
	Wscript.Echo "���: "& CStr(Err.Number) & vbNewLine & Err.Description & vbNewLine & "������ ��� �������� ������."
	DispErr Err.Number, Err.Description
End If

' ������� ����� ��� ������ ������� � ������� �����
name="����� ��� ������ ������� �� ����"
fPath = oShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\" & name
Set fso=WScript.CreateObject("Scripting.FileSystemObject") 
' ���� ����� �� ����������, �� ������� �����
if Not fso.FolderExists(fPath) then fso.CreateFolder(fPath)

dim objSD, objACE
Const FILE_SHARE = 0
Const MAXIMUM_CONNECTIONS = 25
Const ACCESS = 1245631 '����� �� ������ � ��������� � ����������� ������ �������
strComputer = "."
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objNewShare = objWMI.Get("Win32_Share")
errReturn = objNewShare.Create (fPath, name, FILE_SHARE, MAXIMUM_CONNECTIONS)
if errReturn=0 then
	set objSecSettings = objWMI.Get("Win32_LogicalShareSecuritySetting.Name='" & name & "'")
	If objSecSettings.GetSecurityDescriptor(objSD) = 0 Then
            If Not IsNull(objSD.DACL) Then
                For Each objACE In objSD.DACL
                    objACE.AccessMask = ACCESS
                Next
                Set objACE = Nothing
                errReturn = objSecSettings.SetSecurityDescriptor(objSD)
                Select Case errReturn
                    Case 0: errDescription = "�������� ����������. ����� ���������"
                    Case 2: errDescription = "����������� ������ � ����������� ����������." 
						DispErr errReturn, errDescription
                    Case 9: errDescription = "��� ���������� �������� ������������ ����������."
						DispErr errReturn, errDescription
                    Case 21: errDescription = "������ ������������ �������� ����������."
						DispErr errReturn, errDescription
                    Case Else: errDescription = "����������� ������ � �����: " & errReturn
						DispErr errReturn, errDescription
                End Select
            Else
                errDescription = "������ ���������� �������� � ������� " & UCase(name) & " ����."
				DispErr errReturn, errDescription
            End If
	Else
		errDescription = "�� ������� ��������� ���������� ������������ ������� " & UCase(name)
		DispErr errReturn, errDescription
	End If
	Set objSD = Nothing
    Set objSecSettings = Nothing
ElseIf errReturn=22 then
	errDescription = "������ " & errReturn & ": ����� ������ " & UCase(name) & " ��� ����������"
	DispErr errReturn, errDescription
Else
    errDescription = "������ " & errReturn & " ��� �������� ������� ������ ������� " & UCase(name)
	DispErr errReturn, errDescription
End If
Wscript.Echo errDescription

'���������� ����������� � ���������� NTFS
Dim objWsNet, strDomain, strComputer, strAccount
Dim strPath, xResult, xErr

strAccount = "���"

If StrComp(strAccount, "�������", vbTextCompare) = 0 Then strAccount = "System"
        strPath = fPath
        Set objWsNet = CreateObject("WScript.Network")
        strComputer = objWsNet.ComputerName
        Set objWsNet = Nothing        
        If StrComp(strAccount, "System", vbTextCompare) <> 0 And StrComp(strAccount, "���", vbTextCompare) <> 0 Then
            strDomain = strComputer
        Else
            strDomain = vbNullString
        End If
        xErr = Set_RWEAccess(strDomain, strComputer, strAccount, strPath)
        If IsNumeric(xErr) Then xErr = CStr(xErr)
        Select Case xErr
            Case "-3": xResult = "�� ������� ��������� ��������� ������� ������������ ������ " & UCase(strDomain & "\" & strAccount)
            Case "-2": xResult = "�� ������� ������� ������ ������� " & UCase(strDomain & "\" & strAccount)
            Case "-1": xResult = "�� ������� ��������� ������������ ������������ � ����� " & UCase(strPath)
            Case "0": Wscript.Echo "�������� ����������. ����� NTFS ���������"
            Case "2": xResult = "������ ��������."
            Case "8": xResult = "����������� ������."
            Case "5", "9": xResult = "��� ���������� �������� ������������ ����������."
            Case "21": xResult = "������ ������������ �������� ����������."
            Case Else: WScript.Echo xErr
        End Select
		if xErr <> 0 then 
			Wscript.Echo xErr & ": " & xResult
			DispErr xErr, xResult
		end If

' ������� ���������� ���������� NTFS (������� "������������")
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
                xRes = "������ ���������� �������� (ACL) � ��������� ������� ����."
            End If
        Else
            xRes = "�� ������� ��������� ���������� ������������ �������."
        End If
        Set objSD = Nothing
        Set objSecSettings = Nothing
    Else
        xRes = "������ " & CStr(Err.Number) & vbNewLine & Err.Description
        Err.Clear
    End If
Else
    xRes = "������ " & CStr(Err.Number) & vbNewLine & Err.Description
    Err.Clear
End If
Set objWMI = Nothing
On Error GoTo 0
Set_RWEAccess = xRes
End Function

On Error Resume Next
set oShellLink = oShell.CreateShortcut(oShell.SpecialFolders("Desktop") & "\" & name & ".lnk")
oShellLink.TargetPath = fPath
if Err.Number=0 then
	oShellLink.Save
	Wscript.Echo "����� �� ����� ��� ������ ������� �� ����"
Else
	Wscript.Echo "���: "& CStr(Err.Number) & vbNewLine & Err.Description & vbNewLine & "������ ��� �������� ������."
	DispErr errReturn, errDescription
End If
Wscript.Echo "������ �������� ������"
'if NoErrors=true then 
'	oShell.AppActivate "Command Prompt"
'	oShell.SendKeys "color 20~ pause~ exit~"
'End If

Wscript.Sleep(5000)

	