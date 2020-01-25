
Class clsRegistry
    BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Public Enum RegKeys
        HKEY_CURRENT_CONFIG = &H80000005
        HKEY_CURRENT_USER = &H80000001
        HKEY_DYN_DATA = &H80000006
        HKEY_LOCAL_MACHINE = &H80000002
        HKEY_PERF_ROOT = HKEY_LOCAL_MACHINE
        HKEY_PERFORMANCE_DATA = &H80000004
        HKEY_USERS = &H80000003
        HKEY_CLASSES_ROOT = &H80000000
    End Enum

    ' intern
    Private Const KEY_QUERY_VALUE = &H1
    Private Const KEY_SET_VALUE = &H2
    Private Const KEY_CREATE_SUB_KEY = &H4
    Private Const KEY_ENUMERATE_SUB_KEYS = &H8
    Private Const KEY_NOTIFY = &H10
    Private Const KEY_CREATE_LINK = &H20
    Private Const KEY_READ = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
    Private Const KEY_WRITE = KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
    Private Const KEY_EXECUTE = KEY_READ
    Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK
    Private Const ERROR_SUCCESS = 0&

    Public Enum RegValueTypes
        REG_NONE = 0                  ' No value type
        REG_SZ = 1                    ' Unicode nul terminated string
        REG_EXPAND_SZ = 2             ' Unicode nul terminated string (with environment variable references)
        REG_BINARY = 3                ' Free form binary
        REG_DWORD = 4                 ' 32-bit number
        REG_DWORD_LITTLE_ENDIAN = 4   ' 32-bit number (same as REG_DWORD)
        REG_DWORD_BIG_ENDIAN = 5      ' 32-bit number
        REG_LINK = 6                  ' Symbolic Link (unicode)
        REG_MULTI_SZ = 7              ' Multiple Unicode strings
    End Enum

    Private Const REG_OPTION_NON_VOLATILE = &H0
    Private Const REG_CREATED_NEW_KEY = &H1

    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Any) As Long
    Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
    Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Private Declare Function RegSetValueEx_String Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
    Private Declare Function RegSetValueEx_DWord Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
    Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
    Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Byte, ByRef lpcbData As Long) As Long
    Dim root As RegKeys

    Private Function DeleteKeyLoop(ByVal sKey$) As Boolean
        On Error Resume Next
        Dim lResult&, cSubKeys As Collection, i%
        If ListSubKeys(sKey, cSubKeys) Then
            For i = 1 To cSubKeys.Count
                DeleteKeyLoop Replace(sKey & "\" & cSubKeys(i), "\\", "\")
            Next i
        End If
        lResult = RegDeleteKey(root, sKey)
        DeleteKeyLoop = (lResult = ERROR_SUCCESS)
    End Function

    Public Function ExistKey(ByVal sKey$) As Boolean
        'checks if a key exists
        Dim lResult&, keyhandle&
        lResult = RegOpenKeyEx(root, sKey, 0, KEY_READ, keyhandle)
        If lResult = ERROR_SUCCESS Then RegCloseKey keyhandle
        ExistKey = (lResult = ERROR_SUCCESS)
    End Function

    Public Function GetValueType(ByVal sKey$, ByVal sField$) As RegValueTypes
        'gets the type of a value
        Dim lResult&, keyhandle&, dwType&, zw&, puffersize&, puffer$
        lResult = RegOpenKeyEx(root, sKey, 0, KEY_READ, keyhandle)
lResult = RegQueryValueEx(keyhandle, sField, 0&, dwType, ByVal 0&, puffersize)
        If lResult = ERROR_SUCCESS Then
            GetValueType = dwType
        End If
    End Function

    Public Function GetValue(ByVal sKey$, ByVal sField$, ByRef sValue$) As Boolean
        'gets a value from the registry
        On Error Resume Next
        Dim lResult&, keyhandle&, dwType&, zw&, puffersize&, puffer$
        lResult = RegOpenKeyEx(root, sKey, 0, KEY_READ, keyhandle)
        GetValue = (lResult = ERROR_SUCCESS)
        If lResult <> ERROR_SUCCESS Then Exit Function ' key doesn't exist
lResult = RegQueryValueEx(keyhandle, sField, 0&, dwType, ByVal 0&, puffersize)
        GetValue = (lResult = ERROR_SUCCESS)
        If lResult <> ERROR_SUCCESS Then Exit Function ' field doesn't exist
        Select Case dwType
            Case REG_SZ, REG_EXPAND_SZ   ' null-terminated String
                puffer = Space$(puffersize + 1)
    lResult = RegQueryValueEx(keyhandle, sField, 0&, dwType, ByVal puffer, puffersize)
                GetValue = (lResult = ERROR_SUCCESS)
                If lResult <> ERROR_SUCCESS Then Exit Function ' Error while reading the value
                sValue = puffer
            Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN   ' 32-Bit Number   !!!! Word
                puffersize = 4      ' = 32 Bit
                lResult = RegQueryValueEx(keyhandle, sField, 0&, dwType, zw, puffersize)
                GetValue = (lResult = ERROR_SUCCESS)
                If lResult <> ERROR_SUCCESS Then Exit Function ' Error while reading the value
                sValue = zw
        End Select
        If lResult = ERROR_SUCCESS Then RegCloseKey keyhandle
        GetValue = True
        If Asc(Right$(sValue, 1)) = 32 Then sValue = Left$(sValue, Len(sValue) - 2)
    End Function

    Public Function CreateKey(ByVal sKey$) As Boolean
        'creates a key
        Dim lResult&, keyhandle&, Action&
        lResult = RegCreateKeyEx(root, sKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, keyhandle, Action)
        If lResult = ERROR_SUCCESS Then
            If RegFlushKey(keyhandle) = ERROR_SUCCESS Then RegCloseKey keyhandle
        Else
            CreateKey = False
            Exit Function
        End If
        CreateKey = (Action = REG_CREATED_NEW_KEY)
    End Function

    Public Function SetValue(ByVal sKey$, ByVal sField$, ByVal Value As Object, Optional ByVal RegValueType As RegValueTypes = REG_SZ) As Boolean
        On Error Resume Next
        Dim lResult&, keyhandle&, S$, l&, ValueLong&
        lResult = RegOpenKeyEx(root, sKey, 0, KEY_ALL_ACCESS, keyhandle)
        If lResult <> ERROR_SUCCESS Then
            SetValue = False
            Exit Function
        End If
        Select Case RegValueType
            Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ, REG_BINARY
                lResult = RegSetValueEx_String(keyhandle, sField, 0, RegValueType, CStr(Value), Len(CStr(Value)) + 1)
            Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
                ValueLong = CLng(Value)
                lResult = RegSetValueEx_DWord(keyhandle, sField, 0, RegValueType, ValueLong, 4)
        End Select
        RegCloseKey keyhandle
        SetValue = (lResult = ERROR_SUCCESS)
    End Function

    Public Function DeleteKey(ByVal sKey$, Optional ByVal bDeleteSubKeys As Boolean = False) As Boolean
        'deletes a key
        Dim lResult&
        'prevent from very worse delete
        If sKey = "" Then Exit Function
        If bDeleteSubKeys Then
            DeleteKey = DeleteKeyLoop(sKey)
        Else
            lResult = RegDeleteKey(root, sKey)
            DeleteKey = (lResult = ERROR_SUCCESS)
        End If
    End Function

    Public Function DeleteValue(ByVal sKey$, ByVal sField$) As Boolean
        'deletes a value
        Dim lResult&, keyhandle&
        lResult = RegOpenKeyEx(root, sKey, 0, KEY_ALL_ACCESS, keyhandle)
        If lResult <> ERROR_SUCCESS Then
            DeleteValue = False
            Exit Function
        End If
        lResult = RegDeleteValue(keyhandle, sField)
        DeleteValue = (lResult = ERROR_SUCCESS)
        RegCloseKey keyhandle
    End Function

    Public Function GetValueDirect(ByVal sKey$, ByVal sField$) As String
        'gets a value and returns it directly
        Dim sRetStr$
        If GetValue(sKey$, sField$, sRetStr) Then
            GetValueDirect = sRetStr
        End If
    End Function

    Public Function ListSubKeys(ByVal sKey$, ByRef cSubKeys As Collection) As Boolean
        'returns all subkeys of a key
        On Error Resume Next
        Dim hKey&, Result&, Cnt&
        Dim sSubKey$, lSubKeyLen&
        Result = RegOpenKeyEx(root, sKey, 0, KEY_READ, hKey)
        If Result = ERROR_SUCCESS Then
            cSubKeys = New Collection
            Do
                lSubKeyLen = 256
                sSubKey = Space(lSubKeyLen)
                Result = RegEnumKey(hKey, Cnt, sSubKey, lSubKeyLen)
                If Result = ERROR_SUCCESS Then
                    sSubKey = Left(sSubKey, InStr(1, sSubKey, Chr(0)) - 1)
                    cSubKeys.Add sSubKey
                End If
                Cnt = Cnt + 1
            Loop Until Result <> ERROR_SUCCESS
            ListSubKeys = True
        End If
    End Function

    Public Function ListValues(ByVal sKey$, ByRef cValues As Collection) As Boolean
        'returns all subkeys of a key
        On Error Resume Next
        Dim hKey&, Result&, Cnt&, lTemp&, sBuffer As Byte
        Dim sSubKey$, lSubKeyLen&, lType&
        Result = RegOpenKeyEx(root, sKey, 0, KEY_READ, hKey)
        If Result = ERROR_SUCCESS Then
            cValues = New Collection
            Do
                lSubKeyLen = 256
                sSubKey = Space(lSubKeyLen)
        Result = RegEnumValue(hKey, Cnt, sSubKey, lSubKeyLen, 0, lType, ByVal 0&, lTemp)
                If Result = ERROR_SUCCESS Then
                    sSubKey = Left(sSubKey, InStr(1, sSubKey, Chr(0)) - 1)
                    cValues.Add sSubKey
                End If
                Cnt = Cnt + 1
            Loop Until Result <> ERROR_SUCCESS
            ListValues = True
        End If


    End Function

Public Property Get RootKey() As RegKeys
RootKey = root
End Property

Public Property Let RootKey(ByVal vNewValue As RegKeys)
root = vNewValue
End Property

    Private Sub Class_Initialize()
        root = HKEY_CURRENT_USER
    End Sub
End Class
