Attribute VB_Name = "modRegUtils"
'***************************************************************
'Windows API/Global Declarations for :cReadWriteEasyReg (Updated
'     Again)
'***************************************************************
Option Explicit
Option Base 1



Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

' Reg Key Security Options...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL


' Reg Key ROOT Types...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' Reg Create Type Values...
Public Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Public Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Public Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore


Public Const ERROR_SUCCESS = 0                  ' Return Value...
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number


Public Const gREGKEY_APPROOT = "SOFTWARE\SCI Custom Services\Broadcaster" ' ScreenSaver registry subkey
Public Const qREGKEY_OPTIONSROOT = "\Options"
Public Const qREGKEY_COMPUTERSROOT = "\Computers"

Public Const RS_OPTIONS = "Options"
Public Const RK_DATE = "UpdateDate"
Public Const RK_TIME = "UpdateTime"
Public Const RK_OCCURANCE = "Occurance"
Public Const RK_UDLNAME = "UDLName"
'Public Const RK_COMPANY = "Company"
Public Const RK_SERVER = "Server"
Public Const RK_TABLE = "Table"

Public Const RS_COMPUTERS = "Computers"
Public Const RK_COMPUTERS = "Computers"
Public Const RV_COMPUTER = "Computer"
Public Const RK_COMPUTERCOUNT = "Count"



Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpKeyName As String) As Long
Public Function KeyExists(KeyRoot As Long, KeyName As String) As Boolean
Dim rslt As Long
Dim hKey As Long
  rslt = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
  KeyExists = (rslt = ERROR_SUCCESS)
  rslt = RegCloseKey(hKey)
End Function
'------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'------------------------------------------------------------
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                               ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
'------------------------------------------------------------
GetKeyError:    ' Cleanup After An Error Has Occured...
'------------------------------------------------------------
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
'------------------------------------------------------------
End Function
'------------------------------------------------------------

'------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
'------------------------------------------------------------
    Dim rc As Long                                      ' Return Code
    Dim hKey As Long                                    ' Handle To A Registry Key
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type
'------------------------------------------------------------
    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- Create/Open Registry Key...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' Create/Open //KeyRoot//KeyName

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Errors...

    '------------------------------------------------------------
    '- Create/Modify Key Value...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' A Space Is Needed For RegSetValueEx() To Work...

    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, Len(SubKeyValue))   ' Create/Modify Key Value

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Error
    '------------------------------------------------------------
    '- Close Registry Key...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' Close Key

    UpdateKey = True                                    ' Return Success
    Exit Function                                       ' Exit
'------------------------------------------------------------
CreateKeyError:
'------------------------------------------------------------
    UpdateKey = False                                   ' Set Error Return Code
    rc = RegCloseKey(hKey)                              ' Attempt To Close Key
'------------------------------------------------------------
End Function
'------------------------------------------------------------

'------------------------------------------------------------
Public Function DeleteKey(KeyRoot As Long, KeyName As String, SubKeyName As String) As Boolean
'------------------------------------------------------------
    Dim rc As Long                                      ' Return Code
    Dim hKey As Long                                    ' Handle To A Registry Key
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type
'------------------------------------------------------------
    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- Create/Open Registry Key...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' Create/Open //KeyRoot//KeyName

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Errors...

    rc = RegDeleteKey(hKey, SubKeyName)

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Error
    '------------------------------------------------------------
    '- Close Registry Key...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' Close Key

    DeleteKey = True                                    ' Return Success
    Exit Function                                       ' Exit
'------------------------------------------------------------
CreateKeyError:
'------------------------------------------------------------
    DeleteKey = False                                   ' Set Error Return Code
    rc = RegCloseKey(hKey)                              ' Attempt To Close Key
'------------------------------------------------------------
End Function
'------------------------------------------------------------

Public Sub LoadSettings()
Dim i As Integer

'------------------------------------------------------------
    Dim RegVal As String
'------------------------------------------------------------
'    ' Get Update Date
'    RegVal = ""
'    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_DATE, RegVal)
'    UpdateDate = RegVal
'    If UpdateDate = "" Then
'      UpdateDate = CStr(Date)
'    End If
'
'    ' Get Update Time
'    RegVal = ""
'    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_TIME, RegVal)
'    UpdateTime = RegVal
'    If UpdateTime = "" Then
'      UpdateTime = "4:00 AM"
'    End If
'
'    ' Get Copy Occurance
'    RegVal = ""
'    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_OCCURANCE, RegVal)
'    Occurance = Val(RegVal)
'    If Occurance = 0 Then
'      Occurance = oDaily
'    End If
'
'    UpdateDateTime = CDate(UpdateDate & " " & UpdateTime)
'
'    ' Get UDL File Name
'    RegVal = ""
'    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_UDLNAME, RegVal)
'    UdlPathName = RegVal
'    If UdlPathName = "" Then
'      UdlPathName = App.Path & "\" & "exch.udl"
'    End If
'
'    ' Get Server Name
'    RegVal = ""
'    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_SERVER, RegVal)
'    ServerName = RegVal
'    If ServerName = "" Then
'      ServerName = DEF_SERVER
'    End If
'
'
'    ' Get Table Name to copy to
'    RegVal = ""
'    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_TABLE, RegVal)
'    TableName = RegVal
'    If TableName = "" Then
'      TableName = DEF_TABLE
'    End If

    ' Get the computer count
    RegVal = ""
    Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_COMPUTERSROOT, RK_COMPUTERCOUNT, RegVal)
    ComputerCount = Val(RegVal)
   If ComputerCount <> "" Then
      ComputerCount = CInt(ComputerCount)
      If ComputerCount <> 0 Then
        ReDim arrComputers(ComputerCount) As CompRec
        For i = 1 To ComputerCount
          RegVal = ""
          Call GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_COMPUTERSROOT, RV_COMPUTER & i, RegVal)
          arrComputers(i).netName = Mid(RegVal, 1, InStr(RegVal, "-") - 1)
          arrComputers(i).friendName = Mid(RegVal, InStr(RegVal, "-") + 1)
        Next i
      End If
   Else
     ComputerCount = 0
   End If
End Sub

Public Sub SaveSettings()
Dim i As Integer
Dim RegVal As String                                ' String value of registry key
Dim rslt As Boolean

' RegVal = CStr(UpdateDate)
' Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_DATE, RegVal)
'
' RegVal = CStr(UpdateTime)
' Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_TIME, RegVal)
'
' RegVal = CStr(Occurance)
' Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_OCCURANCE, RegVal)
'
' RegVal = CStr(UdlPathName)
' Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_UDLNAME, RegVal)
'
' RegVal = CStr(ServerName)
' Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_SERVER, RegVal)
'
' RegVal = CStr(TableName)
' Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_OPTIONSROOT, RK_TABLE, RegVal)
 
 'save the computers
 rslt = True
 If KeyExists(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_COMPUTERSROOT) Then
   rslt = DeleteKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT, RK_COMPUTERS)
 End If
 If rslt Then
   'save the company count first
   RegVal = CStr(ComputerCount)
   Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_COMPUTERSROOT, RK_COMPUTERCOUNT, RegVal)
   If ComputerCount <> 0 Then
     For i = 1 To ComputerCount
       RegVal = arrComputers(i).netName & "-" & arrComputers(i).friendName
       Call UpdateKey(HKEY_LOCAL_MACHINE, gREGKEY_APPROOT & qREGKEY_COMPUTERSROOT, RV_COMPUTER & i, RegVal)
     Next i
   End If
 Else
 
 End If
End Sub








