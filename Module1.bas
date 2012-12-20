Attribute VB_Name = "Module1"
Option Explicit

Public Enum RegType
    REG_SZ = 1
    REG_DWORD = 4
End Enum

Public Enum BaseKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_QUERY_VALUE = &H1


Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long



Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As BaseKeys)
'Purpose: create a new regestry key
'Parameters: sNewKeyName - name of regestry key to be created, string
'            lPredefinedKey - location to create new key in regestry, BaseKeys
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function
1    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
        vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
2    RegCloseKey (hNewKey)
End Sub

Sub SetKeyValue(lPredefinedKey As BaseKeys, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As RegType)
'Purpose: set value of existing key
'Parameters: lPredefinedKey - location of registry key, BaseKeys
'            sKeyName - name of key to set value in, string
'            sValueName - name of regestry value to be set, string
'            vValueSetting - value to be set, variant
'            lValueType - type of value to set into registry, RegType
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key
    'open the specified key
1    lRetVal = RegCreateKeyEx(lPredefinedKey, sKeyName, 0&, _
        vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
'     lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
2    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
3    RegCloseKey (hKey)
End Sub

Function QueryValue(lPredefinedKey As BaseKeys, sKeyName As String, sValueName As String, Optional sDefaultValue As Variant) As Variant
'Purpose: rerieve value from regestry
'Parameters: lPredefinedKey - location of registry key, BaseKeys
'            sKeyName - name of key to set value in, string
'            sValueName - name of regestry value to be set, string
'            sDefaultValue - if bad value or no value at key this is returned, variant, optional
    On Error GoTo ErrorHandler
    If IsMissing(sDefaultValue) Then sDefaultValue = ""
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value
    
1    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    'lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    'lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey)
2    lRetVal = QueryValueEx(hKey, sValueName, vValue)
3    If lRetVal = ERROR_BADKEY Then QueryValue = sDefaultValue
4    RegCloseKey (hKey)
5    If IsEmpty(vValue) Then
6        QueryValue = sDefaultValue
7    Else
8        QueryValue = vValue
    End If
    Exit Function
ErrorHandler:
    QueryValue = sDefaultValue
End Function



'SetValueEx and QueryValueEx Wrapper Functions:
Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
'Purpose: set a value to the registry
'Parameters: hKey - registry key to set, long
'            sValueName - name of registry value to set, string
'            lType - type used in registry value, long
'            vValue - value to be place in regestry, variant
    Dim lValue As Long
    Dim sValue As String
1    Select Case lType
        Case REG_SZ
3            sValue = vValue & Chr$(0)
4            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
6            lValue = vValue
7            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
'Purpose: get a value from the registry
'Parameters: lhKey - regestry key to get, long
'            szValueName - name of registry value to get, string
'            vValue - registry value to get, value
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    
'    On Error GoTo QueryValueExError
    
    ' Determine the size and type of data to be read
1    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
'    If lrc <> ERROR_NONE Then Error 5
    
2    Select Case lType
        ' For strings
        Case REG_SZ:
4            sValue = String(cch, 0)
5            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
6            If lrc = ERROR_NONE Then
7                vValue = Left$(sValue, cch - 1)
8            Else
9                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD:
11            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
12            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
        'all other data types not supported
14            lrc = -1
    End Select
    
QueryValueExExit:
15    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit
End Function

Function DeleteKey(lPredefinedKey As BaseKeys, strKey As String)
    RegDeleteKey lPredefinedKey, strKey
End Function

Function DeleteValue(lPredefinedKey As BaseKeys, strKey As String, strVal As String)
    Dim lRetVal, hKey As Long
    lRetVal = RegOpenKeyEx(lPredefinedKey, strKey, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegDeleteValue(hKey, strVal)
    RegCloseKey (hKey)
End Function
''


