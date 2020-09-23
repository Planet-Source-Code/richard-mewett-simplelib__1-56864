Attribute VB_Name = "SimpleLib"
Option Explicit

'#####################################################################################
'Windows API Declarations / Types / Constants
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpvbDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type BROWSEINFO
    hOwner           As Long
    pidlRoot         As Long
    pszDisplayName   As String
    lpszTitle        As String
    ulFlags          As Long
    lpfn             As Long
    lParam           As Long
    iImage           As Long
End Type
'#####################################################################################

'Used to specify the filter type for the FilterKey function
Public Enum FilterKey
    dt_Integer = 1
    dt_Float = 2
    dt_UCase = 3
    dt_LCase = 4
    dt_Date = 5
    dt_Time = 6
    dt_Text = 7
    dt_Phone = 8
End Enum

Public Function FileExists(sFileName As String) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFileName, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FileExists = False
        Else
            FileExists = True
        End If
    End If
End Function

Public Function FilterKey(ByVal KeyAscii As Integer, ByVal nMask As FilterKey) As Integer
    '################################################################################
    'Purpose: Simple Keyboard filter (i.e. Make TextBox accept only Numeric input)
    '################################################################################
    
    If KeyAscii <> vbKeyBack Then
        Select Case nMask
        Case dt_Integer, dt_Float
            Select Case KeyAscii
            Case 45 '-
            Case 46 '.
                If nMask = dt_Integer Then
                    KeyAscii = 0
                End If
            Case 48 To 57 '0-9
            
            Case Else
                KeyAscii = 0
            End Select
        
        Case dt_Date
            Select Case KeyAscii
            Case 48 To 57 '0 - 9
            Case 46, 47 '. /
            Case Else
                KeyAscii = 0
            End Select
        
        Case dt_Time
            Select Case KeyAscii
            Case 48 To 58 '0 - 9 + :
            Case 46 '.
            Case Else
                KeyAscii = 0
            End Select
        
        Case dt_Text
            Select Case KeyAscii
            Case 48 To 57 '0 - 9
                KeyAscii = 0
            End Select
            
        Case dt_Phone
            Select Case KeyAscii
            Case 32, 40, 41 ' ( )
            Case 48 To 57 '0-9
            
            Case Else
                KeyAscii = 0
            End Select
        
        Case dt_LCase
            KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
        Case dt_UCase
            KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
        End Select
    End If
    
    FilterKey = KeyAscii
End Function

Public Function FolderExists(sFolder As String) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function

Public Function GetIniBool(ByVal sSection As String, ByVal sEntry As String, bDefault As Boolean, ByVal sFileName As String) As Boolean
    '################################################################################
    'Purpose: Get a Boolean parameter from Ini file
    '################################################################################

    Dim sBuffer As String * 100
    Dim sData As String
    Dim lLength As Long
    
    lLength = GetPrivateProfileString(sSection, sEntry, "", sBuffer, 100, sFileName)
    sData = Trim$(Left$(sBuffer, lLength))

    Select Case LCase$(sData)
    Case "1", "true", "yes"
        GetIniBool = True
    Case "0", "false", "no"
        GetIniBool = False
    Case Else
        If bDefault Then
            PutIniBool sSection, sEntry, bDefault, sFileName
        End If
        GetIniBool = bDefault
    End Select
End Function

Public Function GetIniDbl(ByVal sSection As String, ByVal sEntry As String, ByVal dDefault As Double, ByVal sFileName As String) As Double
    '################################################################################
    'Purpose: Get a Double parameter from Ini file
    '################################################################################

    Dim sBuffer As String * 100
    Dim sData As String
    Dim lLength As Long
    
    On Local Error Resume Next
    
    lLength = GetPrivateProfileString(sSection, sEntry, "", sBuffer, 100, sFileName)
    sData = Left$(sBuffer, lLength)
    If Len(sData) = 0 Then
        sData = CStr(dDefault)
        lLength = WritePrivateProfileString(sSection, sEntry, sData, sFileName)
    End If

    GetIniDbl = Val(sData)
End Function

Public Function GetIniLng(ByVal sSection As String, ByVal sEntry As String, ByVal lDefault As Long, ByVal sFileName As String) As Long
    '################################################################################
    'Purpose: Get a Long parameter from Ini file
    '################################################################################

    Dim sBuffer As String * 100
    Dim sData As String
    Dim lLength As Long
    
    On Local Error Resume Next
    
    lLength = GetPrivateProfileString(sSection, sEntry, "", sBuffer, 100, sFileName)
    sData = Left$(sBuffer, lLength)
    If Len(sData) = 0 Then
        sData = CStr(lDefault)
        lLength = WritePrivateProfileString(sSection, sEntry, sData, sFileName)
    End If

    GetIniLng = Val(sData)
End Function


Public Function GetIniStr(ByVal sSection As String, ByVal sEntry As String, ByVal sDefault As String, ByVal sFileName As String) As String
    '################################################################################
    'Purpose: Get a String parameter from Ini file
    '################################################################################

    Dim sBuffer As String * 100
    Dim sData As String
    Dim lLength As Long
    
    lLength = GetPrivateProfileString(sSection, sEntry, "", sBuffer, 100, sFileName)
    sData = Left$(sBuffer, lLength)
    If Len(sData) = 0 And sDefault <> "" Then
        sData = sDefault
        lLength = WritePrivateProfileString(sSection, sEntry, sDefault, sFileName)
    End If

    GetIniStr = sData
End Function

Public Sub HighlightText(ctlControl As Control, Optional SetFocus As Boolean)
    With ctlControl
        .SelStart = 0
        .SelLength = Len(.Text)
        
        If SetFocus And .Enabled And .Visible Then
            .SetFocus
        End If
    End With
End Sub

Public Sub PutIniBool(ByVal sSection As String, ByVal sEntry As String, ByVal bValue As Integer, ByVal sFileName As String)
    '################################################################################
    'Purpose: Put a Boolean value into an Ini file
    '################################################################################

    Dim lTemp As Long
    Dim sValue As String

    If bValue Or bValue = 1 Then
        sValue = "Yes"
    Else
        sValue = "No"
    End If
    
    lTemp = WritePrivateProfileString(sSection, sEntry, sValue, sFileName)
End Sub

Public Sub PutIniDbl(ByVal sSection As String, ByVal sEntry As String, ByVal dValue As Double, ByVal sFileName As String)
    '################################################################################
    'Purpose: Put a Double value into an Ini file
    '################################################################################

    Dim lTemp As Long
    
    lTemp = WritePrivateProfileString(sSection, sEntry, Format$(dValue), sFileName)
End Sub

Public Sub PutIniLng(ByVal sSection As String, ByVal sEntry As String, ByVal lValue As Long, ByVal sFileName As String)
    '################################################################################
    'Purpose: Put a Long value into an Ini file
    '################################################################################

    Dim lTemp As Long
    
    lTemp = WritePrivateProfileString(sSection, sEntry, Format$(lValue), sFileName)
End Sub


Public Sub PutIniStr(ByVal sSection As String, ByVal sEntry As String, ByVal sValue As String, ByVal sFileName As String)
    '################################################################################
    'Purpose: Put a string value into an Ini file
    '################################################################################

    Dim lTemp As Long
    
    lTemp = WritePrivateProfileString(sSection, sEntry, Trim$(sValue), sFileName)
End Sub

Public Function RequestFolder(Owner As Form, Optional Title As String) As String
    '################################################################################
    'Purpose: Show Folder Browser & return selected path
    '################################################################################

    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim sFolder As String
    Dim pos As Integer
    
    'This set the parameters required for the Windows system call
    With bi
        'hwnd of the window (Form) that messages are directed to
        .hOwner = Owner.hWnd
        
        'Setting to NULL starts the Browser at the desktop folder
        .pidlRoot = 0
    
        'Browser Title
        If Len(Title) = 0 Then
            .lpszTitle = "Select the folder."
        Else
            .lpszTitle = Title
        End If
    
        'This forces the Browser to only show Folders (no files)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
        
    'The call to show the Browser
     pidl = SHBrowseForFolder(bi)
    
    sFolder = Space$(MAX_PATH)
    If SHGetPathFromIDList(pidl, sFolder) Then
        'Find the terminating char in the string to remove unwanted padding
        pos = InStr(sFolder, Chr$(0))
        RequestFolder = Left$(sFolder, pos - 1)
    End If
    
    Call CoTaskMemFree(pidl)
End Function

Public Sub SetComboIndex(cboList As Control, ByVal vValue As Variant, Optional ByVal nDefault As Integer = -1)
    '################################################################################
    'Purpose: Set the ListIndex of a ComboBox/ListBox to the Item matching the
    'Value passed.
    '################################################################################

    Dim lCount As Long
    Dim bFound As Boolean
    Dim bItemData As Boolean
    
    'If we are passed a Long then Search ItemData property instead of Text
    If VarType(vValue) = vbLong Then
        bItemData = True
    End If

    With cboList
        For lCount = 0 To .ListCount - 1
            If bItemData Then
                If vValue = .ItemData(lCount) Then
                    bFound = True
                    .ListIndex = lCount
                    Exit For
                End If
            Else
                If vValue = .List(lCount) Then
                    bFound = True
                    .ListIndex = lCount
                    Exit For
                End If
            End If
        Next lCount
        
        If Not bFound And nDefault >= 0 Then
            .ListIndex = nDefault
        End If
    End With
End Sub

