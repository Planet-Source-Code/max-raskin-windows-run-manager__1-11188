VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Run Manager"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBackup 
      Caption         =   "&Backup..."
      Height          =   285
      Left            =   1590
      TabIndex        =   16
      Top             =   5790
      Width           =   1485
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   6840
      TabIndex        =   15
      Top             =   5460
      Width           =   825
   End
   Begin VB.OptionButton optWinINI 
      Caption         =   "Win.ini Run/Load"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      ToolTipText     =   "Program that are launched from the windows initlization file 'Win.ini'"
      Top             =   60
      Width           =   1575
   End
   Begin VB.OptionButton optStartMenu 
      Caption         =   "StartUp Folder"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      ToolTipText     =   "Program Shortcuts that are launched from the 'Start Menu\Programs\Start Up\' Folder"
      Top             =   60
      Width           =   1395
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete Selection"
      Height          =   285
      Left            =   60
      TabIndex        =   14
      Top             =   5790
      Width           =   1485
   End
   Begin VB.TextBox txtCmdLine 
      Height          =   285
      Left            =   2310
      TabIndex        =   11
      Top             =   5460
      Width           =   4485
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2310
      TabIndex        =   10
      Top             =   5130
      Width           =   5325
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add/Set"
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   5130
      Width           =   915
   End
   Begin VB.ListBox lstCmdLine 
      Height          =   4350
      Left            =   2940
      TabIndex        =   6
      Top             =   720
      Width           =   4725
   End
   Begin VB.ListBox lstName 
      Height          =   4350
      Left            =   60
      TabIndex        =   5
      Top             =   690
      Width           =   2805
   End
   Begin VB.OptionButton optRun2 
      Caption         =   "HKEY_CURRENT_USER Run"
      Height          =   255
      Left            =   2100
      TabIndex        =   2
      ToolTipText     =   "Programs that are found in the 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunServices' Registry Key"
      Top             =   60
      Width           =   2535
   End
   Begin VB.OptionButton optRunServices 
      Caption         =   "Run Services"
      Height          =   255
      Left            =   750
      TabIndex        =   1
      ToolTipText     =   "Service Programs that are found in the 'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices' Registry Key"
      Top             =   60
      Width           =   1305
   End
   Begin VB.OptionButton optRun 
      Caption         =   " Run"
      Height          =   255
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "Programs that are found in the 'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run' Registry Key"
      Top             =   60
      Value           =   -1  'True
      Width           =   675
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command Line:"
      Height          =   195
      Index           =   3
      Left            =   1170
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Name:"
      Height          =   195
      Index           =   2
      Left            =   1170
      TabIndex        =   12
      Top             =   5190
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command Line:"
      Height          =   195
      Index           =   1
      Left            =   3030
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   480
      Width           =   465
   End
   Begin VB.Menu mnuBk 
      Caption         =   "Backup"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "Make HKEY_LOCAL_MACHINE Run Reg File"
      End
      Begin VB.Menu mnu2 
         Caption         =   "Make HKEY_LOCAL_MACHINE RunServices Reg File"
      End
      Begin VB.Menu mnu3 
         Caption         =   "Make HKEY_CURRENT_USER Run Reg File"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Windows Run Manager v1.0

'By Max Raskin, 2 September 2000

'Purpose: Modify Window's StartUp programs in the Registery, Win.INI and the StartUp folder.

Dim hKey As Long, hKey2 As Long, lCount As Long, i As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const LB_SETHORIZONTALEXTENT = (1045)
Dim CurKey As String
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


'Enumerate from HKEY_LOCAL_MACHINE , Run
Private Sub RMEnumRegRun()
    On Error Resume Next
    ClrLists
    hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        lstName.AddItem EnumValue(hKey, i)
        lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
    lstName.ListIndex = 0
    lstCmdLine.ListIndex = 0
End Sub


'Enumerate from HKEY_LOCAL_MACHINE , RunServices
Private Sub RMEnumRegRunServices()
    On Error Resume Next
    ClrLists
    hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        lstName.AddItem EnumValue(hKey, i)
        lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
    lstName.ListIndex = 0
    lstCmdLine.ListIndex = 0
End Sub

'Enumerate from HKEY_CURRENT_USER , Run
Private Sub RMEnumRegRun2()
    On Error Resume Next
    ClrLists
    lstName.ListIndex = lstCmdLine.ListIndex = 1
    hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        lstName.AddItem EnumValue(hKey, i)
        lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
    lstName.ListIndex = 0
    lstCmdLine.ListIndex = 0
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    Dim prvidx As Integer
    prvidx = lstName.ListIndex
    If Trim(txtName.Text) = "" Then
        MsgBox "Enter Name", vbInformation, "No Name"
        txtName.SetFocus
        Exit Sub
    End If
    If Trim(txtCmdLine.Text) = "" Then
        MsgBox "Enter Command Line", vbInformation, "No CmdLine"
        txtCmdLine.SetFocus
    End If
    If optRun.Value = True Then
        If CurKey <> txtName.Text Then
            SetValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", txtName.Text, txtCmdLine.Text
        Else
            SetValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", CurKey, txtCmdLine.Text
        End If
        RMEnumRegRun
    End If
    If optRunServices.Value = True Then
        If CurKey <> txtName.Text Then
            SetValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", txtName.Text, txtCmdLine.Text
        Else
            SetValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", CurKey, txtCmdLine.Text
        End If
        RMEnumRegRunServices
    End If
    If optRun2.Value = True Then
        If CurKey <> txtName.Text Then
            SetValue RegistryKeys.HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", txtName.Text, txtCmdLine.Text
        Else
            SetValue RegistryKeys.HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", CurKey, txtCmdLine.Text
        End If
        RMEnumRegRun2
    End If
    lstCmdLine.ListIndex = prvidx
    lstName.ListIndex = prvidx
End Sub

Private Sub cmdBackup_Click()
    Me.PopupMenu mnuBk
End Sub

Private Sub cmdBrowse_Click()
    Dim r As String
    r = OpenFile
    If r <> "" Then txtCmdLine.Text = r
End Sub

Private Sub cmdDel_Click()
    On Error Resume Next
    Dim prvidx As Integer, msgResult As VbMsgBoxResult
    prvidx = lstName.ListIndex
    If optRun.Value = False Then
        If optRunServices.Value = False Then
            If optRun2.Value = False Then
                Exit Sub
            End If
        End If
    End If
    msgResult = MsgBox("Are you sure you want to delete the program/service '" & lstName.List(prvidx) & "' from the run sequence ?", vbQuestion Or vbYesNo, "Confirm Delete")
    If msgResult = vbNo Then
        Exit Sub
    Else
        'Do Nothing and continue
    End If
    If optRun.Value = True Then
        hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
        DeleteValue hKey, CurKey
        RMEnumRegRun
    End If
    If optRunServices.Value = True Then
        hKey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices")
        DeleteValue hKey, CurKey
        RMEnumRegRunServices
    End If
    If optRun2.Value = True Then
        hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
        DeleteValue hKey, CurKey
        RMEnumRegRun2
    End If
    lstCmdLine.ListIndex = prvidx - 1
    lstName.ListIndex = prvidx - 1
End Sub


Private Sub Form_Load()
    LstAddScroll lstName
    LstAddScroll lstCmdLine
    RMEnumRegRun
End Sub

Private Sub LstAddScroll(Listbox As Listbox)
    SendMessage Listbox.hwnd, LB_SETHORIZONTALEXTENT, 600, 0
End Sub

Private Sub ClrLists()
    lstName.Clear
    lstCmdLine.Clear
End Sub

Private Sub lstCmdLine_Click()
    On Error Resume Next
    lstName.ListIndex = lstCmdLine.ListIndex
    txtCmd.Text = lstCmdLine.List(lstCmdLine.ListIndex)
    CurKey = lstName.List(lstName.ListIndex)
End Sub

Private Sub lstName_Click()
    On Error Resume Next
    lstCmdLine.ListIndex = lstName.ListIndex
    txtName.Text = lstName.List(lstName.ListIndex)
    txtCmdLine.Text = lstCmdLine.List(lstCmdLine.ListIndex)
    CurKey = lstName.List(lstName.ListIndex)
End Sub

Private Sub mnu1_Click()
    optRun_Click
    optRun.SetFocus
    MakeRegFile
End Sub

Private Sub mnu2_Click()
    optRunServices_Click
    optRunServices.SetFocus
    MakeRegFile , 1
End Sub

Private Sub mnu3_Click()
    optRun2_Click
    optRun2.SetFocus
    MakeRegFile 1, 1
End Sub

Private Sub optRun_Click()
    RMEnumRegRun
End Sub

Private Sub optRunServices_Click()
    RMEnumRegRunServices
End Sub

Private Sub optRun2_Click()
    RMEnumRegRun2
End Sub

Private Sub optStartMenu_Click()
    ShellExecute 0, "open", CheckFolderID(StartUp), "", CheckFolderID(StartUp), 1
End Sub

Private Sub optWinINI_Click()
    ShellExecute 0, "open", "notepad.exe", WinDir & "\win.ini", "", 1
End Sub

'Get Windows's Directory
Public Function WinDir() As String
    Dim RetVal As String
    Dim Tmp As String
    Tmp = Space$(255)
    RetVal = GetWindowsDirectory(Tmp, 255)
    WinDir = Trim$(Left$(Tmp, RetVal))
End Function

Private Function MakeRegFile(Optional hKey As Integer, Optional nType As Integer) As String
    On Error Resume Next
    Dim sKey1 As String, sKey2 As String, r As String
    sKey1 = "HKEY_LOCAL_MACHINE"
    sKey2 = "Run"
    If hKey >= 1 Then sKey1 = "HKEY_CURRENT_USER"
    If nType >= 1 Then sKey2 = "RunServices"
    MakeRegFile = "REGEDIT4" & vbCrLf & vbCrLf & "[" & sKey1 & "\Software\Microsoft\Windows\CurrentVersion\" & sKey2 & "]"
    For i = 0 To lstName.ListCount - 1
        MakeRegFile = MakeRegFile & vbCrLf & Chr(34) & lstName.List(i) & Chr(34) & "=" & Chr(34) & lstCmdLine.List(i) & Chr(34)
    Next i
    r = SaveFile
    If r <> "" Then
        Open r For Binary As #1
            Put #1, , MakeRegFile
        Close #1
    End If
End Function
