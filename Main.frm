VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demo of Subs/Functions in Simple.bas library"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHighlight 
      Caption         =   "Highlight Text"
      Height          =   315
      Left            =   6330
      TabIndex        =   18
      Top             =   2550
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheckFile 
      Caption         =   "Check File"
      Height          =   315
      Left            =   5610
      TabIndex        =   16
      Top             =   4110
      Width           =   1395
   End
   Begin VB.ListBox lstFiles 
      Height          =   1620
      Left            =   150
      TabIndex        =   14
      Top             =   4110
      Width           =   5355
   End
   Begin VB.CommandButton cmdLoadScreen 
      Caption         =   "Load Screen Input"
      Height          =   315
      Left            =   2790
      TabIndex        =   13
      Top             =   3330
      Width           =   2505
   End
   Begin VB.CommandButton cmdSaveScreen 
      Caption         =   "Save Screen Input"
      Height          =   315
      Left            =   150
      TabIndex        =   12
      Top             =   3330
      Width           =   2505
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   5130
      TabIndex        =   10
      Top             =   2550
      Width           =   1095
   End
   Begin VB.TextBox txtFolder 
      Height          =   315
      Left            =   150
      TabIndex        =   9
      Top             =   2550
      Width           =   4875
   End
   Begin VB.CommandButton cmdSetIndex 
      Caption         =   "Set To Red"
      Height          =   315
      Left            =   2130
      TabIndex        =   7
      Top             =   1710
      Width           =   1395
   End
   Begin VB.ComboBox cboColour 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1710
      Width           =   1785
   End
   Begin VB.TextBox txtNumeric 
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   870
      Width           =   1395
   End
   Begin VB.TextBox txtUppercase 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   480
      Width           =   1395
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Highlight Sub:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6330
      TabIndex        =   19
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "The last file in the List is an imposter. It does not really exist!"
      Height          =   825
      Left            =   5610
      TabIndex        =   17
      Top             =   4890
      Width           =   1995
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "FileExist Function:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Ini Functions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   11
      Top             =   3060
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "RequestFolder Function:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   2310
      Width           =   2100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "SetComboIndex Sub:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   1410
      Width           =   1785
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "This textbox will only allow Numeric input"
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   930
      Width           =   2865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "FilterKey Function:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   210
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This textbox will only allow UPPERCASE input"
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   540
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    Dim sFolder As String
    
    '****************************************************************************
    'If the user selects Cancel then a zero length ("") string is returned and
    'we ignore it
    '****************************************************************************
    
    sFolder = RequestFolder(Me, "Please select a folder:")
    If Len(sFolder) > 0 Then
        txtFolder.Text = sFolder
    End If
End Sub

Private Sub cmdCheckFile_Click()
    Dim sFile As String
    
    '****************************************************************************
    'FileExists returns True if the File was found
    '****************************************************************************
    
    sFile = lstFiles.Text
    
    If FileExists(sFile) Then '(This is shorthand for If FileExists(sFile) = True Then)
        MsgBox "We found the file!", vbInformation
    Else
        MsgBox "We could not find the file!", vbExclamation
    End If
End Sub

Private Sub cmdHighlight_Click()
    HighlightText txtFolder, True
End Sub

Private Sub cmdLoadScreen_Click()
    '****************************************************************************
    'Read data from our Ini File
    '****************************************************************************
    
    txtUppercase.Text = GetIniStr("ScreenSettings", "UppercaseBox", "", "simple.ini")
    txtNumeric.Text = GetIniDbl("ScreenSettings", "NumericBox", 0, "simple.ini")
    
    cboColour.ListIndex = GetIniLng("ScreenSettings", "ColourIndex", 0, "simple.ini")
    
    txtFolder.Text = GetIniStr("ScreenSettings", "Folder", "", "simple.ini")
End Sub

Private Sub cmdSaveScreen_Click()
    '****************************************************************************
    'Write our data to an Ini File
    '****************************************************************************
    
    PutIniStr "ScreenSettings", "UppercaseBox", txtUppercase.Text, "simple.ini"
    PutIniDbl "ScreenSettings", "NumericBox", txtNumeric.Text, "simple.ini"
    
    PutIniLng "ScreenSettings", "ColourIndex", cboColour.ListIndex, "simple.ini"
    
    PutIniStr "ScreenSettings", "Folder", txtFolder.Text, "simple.ini"
End Sub


Private Sub cmdSetIndex_Click()
    '****************************************************************************
    'Locate Red in the ComboBox
    '****************************************************************************
    
    SetComboIndex cboColour, "Red"
End Sub

Private Sub Form_Load()
    Dim nCount As Integer
    Dim sFile As String
    
    'Put some Colours into the ComboBox
    With cboColour
        .AddItem "Blue"
        .AddItem "Green"
        .AddItem "Red"
        .AddItem "White"
        .AddItem "Yellow"
        .ListIndex = 0
    End With
    
    'Find some files to test FileExist function
    sFile = Dir$("C:\*.*", vbNormal)
    Do While sFile <> ""
        lstFiles.AddItem "C:\" & sFile
        
        'We will stop once we have found 5!
        nCount = nCount + 1
        If nCount = 5 Then
            Exit Do
        End If
        
        sFile = Dir$()
    Loop
    
    lstFiles.AddItem "C:\FileWhichDoesNotExist"
    lstFiles.ListIndex = 0
End Sub

Private Sub txtNumeric_KeyPress(KeyAscii As Integer)
    '****************************************************************************
    'Ensure the char entered is Numeric
    '****************************************************************************
    
    KeyAscii = FilterKey(KeyAscii, dt_Float)
End Sub


Private Sub txtUppercase_KeyPress(KeyAscii As Integer)
    '****************************************************************************
    'Ensure the char entered is UPPERCASE
    '****************************************************************************
    
    KeyAscii = FilterKey(KeyAscii, dt_UCase)
End Sub


