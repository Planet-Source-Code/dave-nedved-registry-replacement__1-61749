Attribute VB_Name = "modReg_rep"
Rem // The following code is code to get reg keys ect...
Rem // uses a few declarations to advapi32.dll for Registry functions.
Rem // PLEASE do NOT get mad at me for not commenting this module
Rem // I wrote this module some time ago... meaning that ad it is my code i couldnt
Rem // realy be stuffed to comment it... :D:D lol
Rem // I should comment more often but it requires much thinking for me...

Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
'--------------------------------------------------
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
'--------------------------------------------------
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub CreateKey(hKey As Long, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)



lRegResult = RegCloseKey(hCurKey)

End Sub

Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegDeleteValue(hCurKey, strValue)
lRegResult = RegCloseKey(hCurKey)

End Sub

Public Function ds_GetRegKey(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

If Not IsEmpty(Default) Then
  ds_GetRegKey = Default
Else
  ds_GetRegKey = ""
End If

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)


  If lValueType = REG_SZ Then
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
    
    intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      ds_GetRegKey = Left$(strBuffer, intZeroPos - 1)
    Else
      ds_GetRegKey = strBuffer
    End If

  End If


lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub ds_SaveRegKey(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))


lRegResult = RegCloseKey(hCurKey)
End Sub


Public Sub SaveSetting(ds_AppName As String, ds_Section As String, ds_key As String, ds_Setting As String)
Dim tmpAppName As String, tmpSection As String, tmpKey As String, tmpSetting As String, tmpStringLine As String

tmpAppName = ds_AppName
tmpSection = ds_Section
tmpKey = ds_key
tmpSetting = ds_Setting
tmpStringLine = tmpAppName & "\" & tmpSection & "\"

ds_SaveRegKey HKEY_LOCAL_MACHINE, "Software\DaTo Software\" & tmpStringLine, ds_key, ds_Setting
End Sub

Public Function GetSetting(ds_AppName As String, ds_Section As String, ds_key As String, Optional ds_Default As String) As String
Dim tmpAppName As String, tmpSection As String, tmpKey As String, tmpDefault As String, tmpStringLine As String

tmpAppName = ds_AppName
tmpSection = ds_Section
tmpKey = ds_key
tmpStringLine = tmpAppName & "\" & tmpSection & "\"

GetSetting = ds_GetRegKey(HKEY_LOCAL_MACHINE, "Software\DaTo Software\" & tmpStringLine, ds_key, ds_Default)
End Function
