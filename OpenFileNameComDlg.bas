Attribute VB_Name = "SaveOpenFileNameComDlg"
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Dim ofn As OPENFILENAME
 
Public Function OpenFile() As String
    Dim lReturnValue As Long
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = frmMain.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*"
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrTitle = "Browse For File To Auto Execute:"
        lReturnValue = GetOpenFileName(ofn)
        If (lReturnValue) Then
                OpenFile = Trim(ofn.lpstrFile)
        End If
 End Function

Public Function SaveFile() As String
    Dim lReturnValue As Long
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = frmMain.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrTitle = "Backup"
    ofn.lpstrFilter = "Registration Files (*.reg)" + Chr$(0) + "*.reg"
    ofn.lpstrDefExt = "*.reg"
    lReturnValue = GetSaveFileName(ofn)
    If (lReturnValue) Then
        SaveFile = Trim(ofn.lpstrFile)
    End If
End Function
