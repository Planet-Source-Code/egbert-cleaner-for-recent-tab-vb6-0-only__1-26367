VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls2000Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////
'-- Class Information
'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////

'-- Author     : Mark Belles
'-- Email      : mrbelles@depcoinc.com
'-- Version    : 1.0
'-- Known bugs : none
   'right now it will only handle the writing of string data types into the registry
   'also, it will only read string and binary data from the registry, i'm working on
   'the countless others, however it's friday, and i want to have some fun!
   'I'll keep working on this, at the start of the week, I'll add in every last function
   'in the api, until it's complete, this is a good start, just to get some feedback
   'i'm also going to make an activeX dll and also a nice handy user control with all
   'of these feature, but, i'll continue to release the source, all feedback is welcome!
   'oh, one more thing, the remote registry, oh ya,,, it's coming!!! give me a coupla days
   'and then it will be completed. Right now, all errors should be contained here in
   'the class, nothing yet has given me any trouble, but let me know.
   '-- laterz

'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////
'-- Start code
'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////

'-- top level registry hive keys
Public Enum eHiveKey
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_CURRENT_USER = &H80000001
   HKEY_DYN_DATA = &H80000006
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_PERFORMANCE_DATA = &H80000004
   HKEY_USERS = &H80000003
End Enum

'-- registry data types
Public Enum eValueType
   REG_BINARY = 3                      ' Free form binary
   REG_DWORD = 4                       ' 32 bit number
   REG_DWORD_BIG_ENDIAN = 5            ' 32-bit number
   REG_DWORD_LITTLE_ENDIAN = 4         ' 32-bit number (same as REG_DWORD)
   REG_EXPAND_SZ = 2                   ' Unicode nul terminated string
   REG_LINK = 6                        ' Symbolic Link (unicode)
   REG_MULTI_SZ = 7                    ' Multiple Unicode strings
   REG_NONE = 0                        ' No value type
   REG_RESOURCE_LIST = 8               ' Resource list in the resource map
   REG_SZ = 1                          ' Unicode nul terminated string
End Enum

'-- key constants
Private Const REG_CREATED_NEW_KEY = &H1          ' New Registry Key created
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Private Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Private Const REG_NOTIFY_CHANGE_LAST_SET = &H4   ' Time stamp
Private Const REG_NOTIFY_CHANGE_NAME = &H1       ' Create or delete (child)
Private Const REG_NOTIFY_CHANGE_SECURITY = &H8
Private Const REG_OPENED_EXISTING_KEY = &H2      ' Existing Key opened
Private Const REG_OPTION_BACKUP_RESTORE = 4      ' open for backup or restore
Private Const REG_OPTION_CREATE_LINK = 2         ' Created key is a symbolic link
Private Const REG_OPTION_NON_VOLATILE = 0        ' Key is preserved when system is rebooted
Private Const REG_OPTION_RESERVED = 0            ' Parameter is reserved
Private Const REG_OPTION_VOLATILE = 1            ' Key is not preserved when system is rebooted
Private Const REG_REFRESH_HIVE = &H2             ' Unwind changes to last flush
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Private Const REG_WHOLE_HIVE_VOLATILE = &H1      ' Restore whole hive volatile
Private Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Private Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

'-- access constants
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_EVENT = &H1
Private Const KEY_NOTIFY = &H10
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

'-- private errors
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

'-- security type
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type

Private Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type

'-- file time type
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

'-- api declares
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKEY As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKEY As Long, phkResult As Long) As Long
'Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKEY As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKEY As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKEY As Long, ByVal lpValueName As String) As Long
'Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKEY As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKEY As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKEY As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKEY As Long) As Long
Private Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKEY As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKEY As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKEY As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKEY As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKEY As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKEY As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKEY As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKEY As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKEY As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKEY As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKEY As Long, ByVal lpSubKey As String) As Long
'-- used specifically for the enumValueNames
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////
'-- private variables
'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////

Private m_hKey As eHiveKey
Private m_lpSubKey As String
Private m_lpValueName As String
Private m_lpValueType As eValueType
'Private m_vValue As Variant

'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////
'-- class constructor and destructor
'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////

Private Sub Class_Initialize()
'-- add code here
End Sub

Private Sub Class_Terminate()
'-- add code here
End Sub

'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////
'-- class events
'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////
'-- private functions and subs
'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////
'-- public functions and subs
'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////

Public Function enumValues(lpValues() As Variant, lpSize As Long, Optional hKEY As eHiveKey, Optional lpSubKey As String) As Boolean
'-- enumerate all of the valuenames of the specified subkeys
'-- return a two dimensional array filled with the valuename, and the associated value
Dim phkResult As Long
Dim lResult As Long
Dim dwIndex As Long
Dim lpValueName As String
Dim lpcbValueName As Long
Dim lpType As Long '-- for the data type returned
Dim lpData(0 To 254) As Byte '-- for the actual value for the valuename
Dim lpcbData As Long
Dim lpBuffer As String
Dim lpHexNibble As String
Dim i As Long

On Error GoTo errHandler
   
   '-- make sure that the array is empty, nothing there
   Erase lpValues
   
   If hKEY = 0 Then
      hKEY = m_hKey
      lpSubKey = m_lpSubKey
   End If
   
   '-- open the specified key
   lResult = RegOpenKeyEx(ByVal hKEY, ByVal lpSubKey, ByVal 0, ByVal KEY_QUERY_VALUE, phkResult)
   
   '-- check the value
   If lResult <> ERROR_SUCCESS Then
      GoTo errHandler
   End If
   
   While lResult = ERROR_SUCCESS
           
      lpValueName = Space$(255)
      lpcbValueName = 255
      lpcbData = 255
      
      '-- enumerate with the specified index
      lResult = RegEnumValue(ByVal phkResult, ByVal dwIndex, ByVal lpValueName, lpcbValueName, ByVal 0, lpType, lpData(0), lpcbData)
      
      '-- check the result
      If lResult = ERROR_SUCCESS Then
      
         ReDim Preserve lpValues(1 To 2, 1 To dwIndex + 1) As Variant
               
         '-- trim up the valuename
         lpValueName = Left$(lpValueName, lpcbValueName)
         
         '-- store the valuename, yes... all u speed freaks, this line could be combined with the line above,
         '-- but jeez, it's late on friday, and i need to see what the heck is actually coming and going
         lpValues(1, dwIndex + 1) = lpValueName
         
         '-- let's just see what we have!... debugging only
'         Debug.Print lpValueName
         
         Select Case lpType
            Case 0 '-- REG_NONE
               '-- don't know how to handle this
               
            Case 1 '-- REG_SZ
               '-- a null terminated string, strip the nullchar
               lpBuffer = Space$(lpcbData - 1)
               CopyMemory ByVal lpBuffer, lpData(0), lpcbData - 1
               lpValues(2, dwIndex + 1) = lpBuffer
               
            Case 2 '-- REG_EXPAND_SZ
               '-- don't know how to handle this
               
            Case 3 '-- REG_BINARY
               lpBuffer = ""
               For i = 0 To lpcbData - 1
                  lpHexNibble = Hex(lpData(i))
                  If Len(lpHexNibble) < 2 Then
                     lpHexNibble = String$(2 - Len(lpHexNibble), "0") & lpHexNibble
                  End If
                  lpBuffer = lpBuffer & lpHexNibble
               Next i
               lpValues(2, dwIndex + 1) = lpBuffer
               
            Case 4 '-- REG_DWORD
               '-- this is a long, I'm really not sure how to deal with this
               lpBuffer = Space$(lpcbData - 1)
               CopyMemory ByVal lpBuffer, lpData(0), lpcbData - 1
               lpValues(2, dwIndex + 1) = lpBuffer
               
            Case 5 '-- REG_DWORD_BIG_ENDIAN
               '-- don't know how to handle this
            Case 6 '-- REG_LINK
               '-- don't know how to handle this
            Case 7 '-- REG_MULTI_SZ
               '-- don't know how to handle this
            Case 8 '-- REG_RESOURCE_LIST
               '-- don't know how to handle this
         End Select
         
         dwIndex = dwIndex + 1
         lpSize = dwIndex

      End If
      
   Wend
   
   enumValues = True
   
Exit Function
closeKey:
   '-- try and close the subkey
   lResult = RegCloseKey(phkResult)
   '-- we've gotten here thru an error, so who cares the return val?
   '-- bail
   GoTo errHandler
   
Exit Function
errHandler:
   enumValues = False
   Erase lpValues

End Function

Public Function enumSubKeys(Optional hKEY As eHiveKey, Optional lpSubKey As String) As Collection
'-- enumerate all of the subkeys of the specified subkey
'-- if the optional params are missing, we'll use the presets
'-- return the enumeration inside a collection... much better than an array, well, not really, I just like collections
Dim phkResult As Long
Dim lResult As Long
Dim lpftLastWriteTime As FILETIME

Dim colSubKeys As New Collection
Dim lpBuffer As String
Dim lBufLen As Long
Dim lpClassName As String
Dim lClassNameLen As Long
Dim dwIndex As Long

On Error GoTo errHandler
      
   If hKEY = 0 Then
      hKEY = m_hKey
      lpSubKey = m_lpSubKey
   End If
   
   lResult = RegOpenKeyEx(ByVal hKEY, ByVal lpSubKey, ByVal 0, ByVal KEY_READ, phkResult)
   
   If lResult = ERROR_SUCCESS Then
      
      While lResult = ERROR_SUCCESS
      
         lpBuffer = Space$(255)
         lBufLen = 255
         lpClassName = Space$(255)
         lClassNameLen = 255
         
         lResult = RegEnumKeyEx(ByVal phkResult, ByVal dwIndex, ByVal lpBuffer, lBufLen, ByVal 0, lpClassName, lClassNameLen, lpftLastWriteTime)
         
         '-- check the result
         If lResult = ERROR_SUCCESS Then
            lpBuffer = Left$(lpBuffer, lBufLen)
            lpClassName = Left$(lpClassName, lClassNameLen)
            colSubKeys.Add lpBuffer
            dwIndex = dwIndex + 1
         End If
      Wend
      
      lResult = RegCloseKey(phkResult)
      If lResult <> ERROR_SUCCESS Then
         GoTo errHandler
      End If
      
   Else
      GoTo errHandler
   End If
   
   '-- made it this far, no errors, all is good
   Set enumSubKeys = colSubKeys
   Set colSubKeys = Nothing
      
Exit Function
closeKey:
   '-- try and close the subkey
   lResult = RegCloseKey(phkResult)
   '-- we've gotten here thru an error, so who cares the return val?
   '-- bail
   GoTo errHandler
   
Exit Function
errHandler:
   Set enumSubKeys = New Collection
   
End Function

Public Function deleteValueName(Optional hKEY As eHiveKey, Optional lpSubKey As String, Optional lpValueName As String) As Boolean
'-- delete the specified valuename from within the specified subkey
'-- check for the optional params, if not, use the presets
Dim phkResult As Long
Dim lResult As Long

On Error GoTo errHandler
   
   If hKEY = 0 Then
      hKEY = m_hKey
      lpSubKey = m_lpSubKey
      lpValueName = m_lpValueName
   End If
   
   '-- open the subkey, so that we have access to the specified valuename
   lResult = RegOpenKeyEx(ByVal hKEY, ByVal lpSubKey, ByVal 0, ByVal KEY_ALL_ACCESS, phkResult)
   
   '-- check the result
   If lResult = ERROR_SUCCESS Then
      
      '-- delete the specified valuename
      lResult = RegDeleteValue(phkResult, lpValueName)
      
      '-- check the result
      If lResult <> ERROR_SUCCESS Then
         GoTo errHandler
      End If
      
      '-- gotta close the key, remember, we have only deleted a "valuename", and not a key
      '-- the key is still open, do not forget what we are doing... lol, ok, that was for me,
      '-- it is late on friday afternoon, and I almost left this dang key open, just me talking...ignore.
      lResult = RegCloseKey(phkResult)
      
      '-- check the result
      If lResult <> ERROR_SUCCESS Then
         GoTo errHandler
      End If
      
   Else
      GoTo errHandler
   End If
   
   '-- made it this far, no errors, all is good
   deleteValueName = True
   
Exit Function
closeKey:
   '-- try and close the subkey
   lResult = RegCloseKey(phkResult)
   '-- we've gotten here thru an error, so who cares the return val?
   '-- bail
   GoTo errHandler
   
Exit Function
errHandler:
   deleteValueName = False
   
End Function

Public Function deleteKey(Optional hKEY As eHiveKey, Optional lpSubKey As String) As Boolean
'-- delete the specified subkey
'-- check for the optional params, if not, use the presets
Dim lResult As Long

On Error GoTo errHandler
   
   If hKEY = 0 Then
      hKEY = m_hKey
      lpSubKey = m_lpSubKey
   End If
   
   '-- delete it
   lResult = RegDeleteKey(ByVal hKEY, ByVal lpSubKey)
   
   '-- check the result
   If lResult <> ERROR_SUCCESS Then
      GoTo errHandler
   End If
   
   '-- made it this far, no errors, all is good
   deleteKey = True
   
Exit Function
errHandler:
   deleteKey = False

End Function

Public Function createKey(Optional hKEY As eHiveKey, Optional lpSubKey As String) As Boolean
'-- create a new subkey, or open an existing key and then close it... that's the way the api works!
'-- check for the optional params, if not, use the presets
Dim phkResult As Long
Dim lResult As Long
Dim lpdwDisposition As Long
Dim lpSecurityAttributes As SECURITY_ATTRIBUTES

On Error GoTo errHandler

   '-- set the security attributes for this handle
   With lpSecurityAttributes
      .nLength = Len(lpSecurityAttributes)
      .lpSecurityDescriptor = 0
      .bInheritHandle = False
   End With
   
   '-- check for the params...
   If hKEY = 0 Then
      hKEY = m_hKey
      lpSubKey = m_lpSubKey
   End If
     
   '-- create the key, if it exists it will be opened, if not it will be created then opened
   '-- lol, either way, it's still gonna get opened.
   
   '-- you can change and use the following line if you want the setting to disappear after the next reboot!
   '-- notice the reg_option... that is the key area here
   'lResult = RegCreateKeyEx(hKey, szSubKey, 0, "", REG_OPTION_VOLATILE, KEY_ALL_ACCESS, lpSecurityAttributes, phkResult,lpdwDisposition)
   
   lResult = RegCreateKeyEx(ByVal hKEY, ByVal lpSubKey, ByVal 0, ByVal "", ByVal REG_OPTION_NON_VOLATILE, ByVal KEY_ALL_ACCESS, lpSecurityAttributes, phkResult, lpdwDisposition)
   
   '-- check the result
   If lResult = ERROR_SUCCESS Then
      
      '-- this is kinda pointless right here. at least in this stage of the class
      '-- but maybe you may want to know what just happened... sometime we'll do something with this
      If lpdwDisposition = REG_CREATED_NEW_KEY Then
         '-- add code here, small bit of debug code ... lol
         '-- Debug.Print "Create a new key [" & lpSubKey & "]->" & Now
      ElseIf lpdwDisposition = REG_OPENED_EXISTING_KEY Then
         '-- add code here
      End If
      
      '-- close the key
      lResult = RegCloseKey(phkResult)
      
      '-- check the result
      If lResult <> ERROR_SUCCESS Then
         GoTo errHandler
      End If
      
   Else
      GoTo errHandler
   End If
   
   '-- made it this far, no errors, all is good
   createKey = True
   
Exit Function
closeKey:
   '-- try and close the subkey
   lResult = RegCloseKey(phkResult)
   '-- we've gotten here thru an error, so who cares the return val?
   '-- bail
   GoTo errHandler
   
Exit Function
errHandler:
   createKey = False

End Function

Public Function setValue(hKEY As eHiveKey, lpSubKey As String, lpValueName As String, vValue As Variant, eType As eValueType) As Boolean
'-- set the value for the specified valuename with the specified data type
'-- if no params are specified, we'll use the other presets
Dim phkResult As Long
Dim lResult As Long
Dim lType As Long
Dim szBuffer As String
Dim lBufLen As Long

On Error GoTo errHandler
   

   lResult = RegOpenKeyEx(ByVal hKEY, ByVal lpSubKey, ByVal 0, ByVal KEY_WRITE, phkResult)
   
   '-- check the result
   If lResult = ERROR_SUCCESS Then
   
'      Select Case eType
'         Case 0
'         Case 1 '-- REG_SZ
'            vValue = vValue & vbNullChar
'         Case 2
'         Case 3
'         Case 4
'         Case 5
'         Case 6
'         Case 7
'         Case 8
'      End Select
      
      szBuffer = vValue
      
      'vValue = vValue & vbNullChar
      lResult = RegSetValueEx(ByVal phkResult, ByVal lpValueName, ByVal 0, ByVal eType, ByVal szBuffer, ByVal Len(vValue))
      
      '-- check the result
      If lResult <> ERROR_SUCCESS Then
         GoTo closeKey
      End If
            
      lResult = RegCloseKey(phkResult)
      
      '-- check the result
      If lResult <> ERROR_SUCCESS Then
         GoTo errHandler
      End If
      
   Else
      GoTo errHandler
   End If
   
   '-- made it this far, no errors, all is good
   setValue = True
   
Exit Function
closeKey:
   '-- try and close the subkey
   lResult = RegCloseKey(phkResult)
   '-- we've gotten here thru an error, so who cares the return val?
   '-- bail
   GoTo errHandler
   
Exit Function
errHandler:
   setValue = False
   
End Function

Function getValue(Optional hKEY As eHiveKey, Optional lpSubKey As String, Optional lpValueName As String) As Variant
'-- return the value in the specified valuename
'-- if no params are specified, we'll use the other presets
Dim phkResult As Long
Dim lResult As Long
Dim hThisKey As eHiveKey
Dim szThisSubKey As String
Dim szThisValueName As String
Dim lpType As Long
Dim lpBuffer As String
Dim lBufLen As Long

On Error GoTo errHandler
   
   '-- check for the optional params
   If hKEY = 0 Then
      hKEY = m_hKey
      lpSubKey = m_lpSubKey
      lpValueName = m_lpValueName
   End If
   
   '-- try and open the subkey
   lResult = RegOpenKeyEx(ByVal hKEY, ByVal lpSubKey, ByVal 0, ByVal KEY_READ, phkResult)
   
   '-- check the result
   If lResult = ERROR_SUCCESS Then
   
      '-- it opened, so let's query the valuename
      '-- prep the buffer
      lpBuffer = Space$(255)
      lBufLen = 255
      
      '-- query the value from the specified valuename...
      lResult = RegQueryValueEx(ByVal phkResult, ByVal lpValueName, ByVal 0, lpType, ByVal lpBuffer, lBufLen)
      
      '-- check the result
      If lResult = ERROR_SUCCESS Then
         
         '-- it opened so let's select which data type it is
         '-- when returning values, i need to do a lot more work to handle all of the various types
         '-- right now it only truely handles strings, binary, and dwords...
         '-- of course those are the big 3, especially if using regedit to create
         Select Case lpType
            Case 0 '-- REG_NONE
               getValue = lpBuffer
            Case 1 '-- REG_SZ
               '-- a null terminated string, strip the nullchar
               getValue = Left$(lpBuffer, lBufLen - 1) '-- <- love how the api tell's us how much we actually got!
            Case 2 '-- REG_EXPAND_SZ = 2
               getValue = lpBuffer
            Case 3 '-- REG_BINARY = 3
               getValue = lpBuffer
            Case 4 '-- REG_DWORD = 4
               getValue = lpBuffer
            Case 5 '-- REG_DWORD_BIG_ENDIAN = 5
               getValue = lpBuffer
            Case 6 '-- REG_LINK = 6
               getValue = lpBuffer
            Case 7 '-- REG_MULTI_SZ = 7
               getValue = lpBuffer
            Case 8 '-- REG_RESOURCE_LIST = 8
               getValue = lpBuffer
         End Select
         
         '-- it opened too, so it's valid, let's close the subkey now
         lResult = RegCloseKey(phkResult)
         
         '-- check the result
         If lResult <> ERROR_SUCCESS Then
            '-- not good, for some reason the key didn't close
            '-- i consider that an error by which to return false
            GoTo errHandler
         End If
                  
      Else
         '-- bail
         GoTo errHandler
      End If
      
   Else
      '-- bail
      GoTo errHandler
   End If
   
Exit Function
closeKey:
   '-- try and close the subkey
   lResult = RegCloseKey(phkResult)
   '-- we've gotten here thru an error, so who cares the return val?
   '-- bail
   GoTo errHandler
   
Exit Function
errHandler:
   getValue = ""

End Function

Public Function getValueNameType(Optional hKEY As eHiveKey, Optional szSubKey As String, Optional szValueName As String) As eValueType
'-- return the data type for the specified valuename
'-- if no valuename is specified, we use the other preset
Dim phkResult As Long
Dim lResult As Long
Dim hThisKey As eHiveKey
Dim szThisSubKey As String
Dim szThisValueName As String
Dim lType As Long
Dim szBuffer As String
Dim lBufLen As Long

On Error GoTo errHandler
   
   '-- check for the optional params
   If hKEY = 0 Then
      hThisKey = m_hKey
      szThisSubKey = m_lpSubKey
      szThisValueName = m_lpValueName
   Else
      hThisKey = hKEY
      szThisSubKey = szSubKey
      szThisValueName = szValueName
   End If
   
   '-- try and open the subkey
   lResult = RegOpenKeyEx(hThisKey, szThisSubKey, 0, KEY_READ, phkResult)
   
   '-- check the result
   If lResult = ERROR_SUCCESS Then
   
      '-- it opened, so let's query the valuename
      '-- prep the buffer
      szBuffer = Space$(255)
      lBufLen = 255
      
      '-- NOTE: VERY IMPORTANT, YOU MUST PASS THE BUFFER BY VALUE (ie. ByVal) TO THIS API
      '-- OR IT WILL/CAN/HAS CRASHED VB6.EXE
      lResult = RegQueryValueEx(phkResult, szThisValueName, 0, lType, ByVal szBuffer, lBufLen)
      
      '-- check the result
      If lResult = ERROR_SUCCESS Then
         
         '-- it opened so let's select which data type it is
         Select Case lType
            Case 0
               getValueNameType = REG_NONE
            Case 1
               getValueNameType = REG_SZ
            Case 2 '-- REG_EXPAND_SZ = 2
               getValueNameType = REG_EXPAND_SZ
            Case 3 '-- REG_BINARY = 3
               getValueNameType = REG_BINARY
            Case 4 '-- REG_DWORD = 4
               getValueNameType = REG_DWORD
            Case 5 '-- REG_DWORD_BIG_ENDIAN = 5
               getValueNameType = REG_DWORD_BIG_ENDIAN
            Case 6 '-- REG_LINK = 6
               getValueNameType = REG_LINK
            Case 7 '-- REG_MULTI_SZ = 7
               getValueNameType = REG_MULTI_SZ
            Case 8 '-- REG_RESOURCE_LIST = 8
               getValueNameType = REG_RESOURCE_LIST
         End Select
         
         '-- it opened too, so it's valid, let's close the subkey now
         lResult = RegCloseKey(phkResult)
         
         '-- check the result
         If lResult <> ERROR_SUCCESS Then
            '-- not good, for some reason the key didn't close
            '-- i consider that an error by which to return false
            GoTo errHandler
         End If
                  
      Else
         '-- bail
         GoTo errHandler
      End If
      
   Else
      '-- bail
      GoTo errHandler
   End If
   
Exit Function
closeKey:
   '-- try and close the subkey
   lResult = RegCloseKey(phkResult)
   '-- we've gotten here thru an error, so who cares the return val?
   '-- bail
   GoTo errHandler
   
Exit Function
errHandler:
   getValueNameType = REG_NONE
   
End Function

Public Function SubKeyExists(Optional hKEY As eHiveKey, Optional szSubKey As String) As Boolean
'-- check for the existence of a subkey
Dim phkResult As Long
Dim lResult As Long
Dim hThisKey As eHiveKey
Dim szThisSubKey As String

On Error GoTo errHandler
   
   If IsMissing(hKEY) Then
      hThisKey = m_hKey
      szThisSubKey = m_lpSubKey
   Else
      hThisKey = hKEY
      szThisSubKey = szSubKey
   End If
   
   '-- attempt to open the key
   lResult = RegOpenKeyEx(hThisKey, szThisSubKey, 0, KEY_READ, phkResult)
   
   '-- check the result
   If lResult = ERROR_SUCCESS Then
      
      '-- it was opened so let's close it
      lResult = RegCloseKey(phkResult)
      
      '-- check the result
      If lResult <> ERROR_SUCCESS Then
         '-- it failed to close
         GoTo errHandler
      End If
   
   Else
   '-- it failed to open, we assume it to not exist
      GoTo errHandler
   End If
        
   '-- made it this far, we were able to successfully open and close the key
   SubKeyExists = True
   
Exit Function
errHandler:
   SubKeyExists = False
   
End Function

Public Function ValueNameExists(Optional hKEY As eHiveKey, Optional szSubKey As String, Optional szValueName As String) As Boolean
'-- check for the existence of a valuename in a subkey
Dim phkResult As Long
Dim lResult As Long
Dim hThisKey As eHiveKey
Dim szThisSubKey As String
Dim szThisValueName As String
Dim lType As Long
Dim szBuffer As String
Dim lBufLen As Long

On Error GoTo errHandler
   
   '-- check for the optional params
   If hKEY = 0 Then
      hThisKey = m_hKey
      szThisSubKey = m_lpSubKey
      szThisValueName = m_lpValueName
   Else
      hThisKey = hKEY
      szThisSubKey = szSubKey
      szThisValueName = szValueName
   End If
   
   '-- try and open the subkey
   lResult = RegOpenKeyEx(hThisKey, szThisSubKey, 0, KEY_READ, phkResult)
   
   '-- check the result
   If lResult = ERROR_SUCCESS Then
      
      '-- it opened, so let's query the valuename
      '-- prep the buffer
      szBuffer = Space$(255)
      lBufLen = 255
      
      '-- NOTE: VERY IMPORTANT, YOU MUST PASS THE BUFFER BY VALUE (ie. ByVal) TO THIS API
      '-- OR IT WILL/CAN/HAS CRASHED VB6.EXE
      lResult = RegQueryValueEx(phkResult, szThisValueName, 0, lType, ByVal szBuffer, lBufLen)
      
      '-- check the result
      If lResult = ERROR_SUCCESS Then
         
         '-- it opened too, so it's valid, let's close the subkey now
         lResult = RegCloseKey(phkResult)
         
         '-- check the result
         If lResult <> ERROR_SUCCESS Then
            '-- not good, for some reason the key didn't close
            '-- i consider that an error by which to return false
            GoTo errHandler
         End If
                  
      Else
         '-- bail
         GoTo errHandler
      End If
      
   Else
      '-- bail
      GoTo errHandler
   End If
   
   '-- made it this far, it exists, and so we return true
   ValueNameExists = True

Exit Function
closeKey:
   '-- try and close the subkey
   lResult = RegCloseKey(phkResult)
   '-- we've gotten here thru an error, so who cares the return val?
   '-- bail
   GoTo errHandler
   
Exit Function
errHandler:
   ValueNameExists = False

End Function

'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////
'-- public properties
'//////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////

Public Property Get RootKey() As eHiveKey
   RootKey = m_hKey
End Property

Public Property Let RootKey(eData As eHiveKey)
   m_hKey = eData
End Property

Public Property Get SubKey() As String
   SubKey = m_lpSubKey
End Property

Public Property Let SubKey(szData As String)
   m_lpSubKey = szData
End Property

Public Property Get ValueName() As String
   ValueName = m_lpValueName
End Property

Public Property Let ValueName(szData As String)
   m_lpValueName = szData
End Property

Public Property Get ValueType() As eValueType
   ValueType = m_lpValueType
End Property

Public Property Let ValueType(eData As eValueType)
   m_lpValueType = eData
End Property

Public Property Get Value() As Variant
   Value = getValue(m_hKey, m_lpSubKey, m_lpValueName)
End Property

Public Property Let Value(vData As Variant)
   Call setValue(m_hKey, m_lpSubKey, m_lpValueName, vData, m_lpValueType)
End Property

