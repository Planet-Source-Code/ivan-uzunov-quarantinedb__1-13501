VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "QuarantineDB"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2925
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2619
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "10.12.2000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "18:08"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2280
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0336
            Key             =   "Sort Ascending"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0448
            Key             =   "Sort Descending"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   "Paste"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditRename 
         Caption         =   "Rename"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertTable 
         Caption         =   "Table"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInsertQuery 
         Caption         =   "Query"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
     
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    On Error Resume Next
    'Just in case if the database is awready closed
    dbDataBase.Close
End Sub

Private Sub mnuEditDelete_Click()
Dim I As Integer
   'Is table or query selected
   If frmDBControl.List1.ListIndex > -1 Then
      I = MsgBox("Are you sure wont to delete " & frmDBControl.List1.List(frmDBControl.List1.ListIndex), vbOKCancel + vbExclamation)
      If I = vbOK Then
         'If the selected is table or query
         If frmDBControl.List1.Tag = "Tables" Then
            dbDataBase.TableDefs.Delete (frmDBControl.List1.List(frmDBControl.List1.ListIndex))
            Call InputTablesToListBox(frmDBControl.List1)
         ElseIf frmDBControl.List1.Tag = "Queries" Then
            dbDataBase.QueryDefs.Delete (frmDBControl.List1.List(frmDBControl.List1.ListIndex))
            Call InputQueriesToListBox(frmDBControl.List1)
         End If
      End If
   Else
      MsgBox "You must select Table or Query", vbCritical, "Quarantine : Error"
   End If
End Sub

Private Sub mnuEditRename_Click()
Dim NewName As String
Dim Count As Integer
Dim OLDTableName As String
Dim OLDQueryName As String
Dim I As Integer
   ' On Error GoTo mnueditrenameerror
    With frmDBControl.List1
         If .ListIndex > -1 Then
            'The new table or query name
            NewName = InputBox("New Name : ", "QuarantineDB Rename " & .List(.ListIndex))
            If .Tag = "Tables" Then
                Count = dbDataBase.TableDefs.Count - 1
                'Find the table and rename it
                For I = 0 To Count
                    OLDTableName = dbDataBase.TableDefs(I).Name
                    If OLDTableName = .List(.ListIndex) Then
                       'Rename the table
                       dbDataBase.TableDefs(I).Name = NewName
                       'and update List1
                       .List(.ListIndex) = NewName
                    End If
                Next I
            ElseIf .Tag = "Queries" Then
                Count = dbDataBase.QueryDefs.Count - 1
                'Find the query and rename it
                For I = 0 To Count
                    OLDQueryName = dbDataBase.QueryDefs(I).Name
                    If OLDQueryName = .List(.ListIndex) Then
                       'Rename the query
                       dbDataBase.QueryDefs(I).Name = NewName
                       'and update List1
                       .List(.ListIndex) = NewName
                    End If
                Next I
            End If
         End If
    End With
End Sub

Private Sub mnuInsertQuery_Click()
   QueryName = InputBox("Input Query Name ", "QuarantineDB")
   If QueryName = "" Then
      MsgBox "Error : Query Name is Empty", vbCritical, "QuarantineDB : Error"
   Else
      frmCreateQuery.Caption = "Create Query : " & QueryName
      frmCreateQuery.Show
   End If
End Sub

Private Sub mnuInsertTable_Click()
   frmCreateTable.Show
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowNewWindow_Click()
    '----
End Sub

Private Sub mnuViewRefresh_Click()
   On Error Resume Next
   ActiveForm.Adodc1.Refresh
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.grdDataGrid.SelText = Clipboard.GetText
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.grdDataGrid.SelText
End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.grdDataGrid.SelText
    ActiveForm.grdDataGrid.SelText = vbNullString
    
End Sub

Private Sub mnuEditUndo_Click()
    MsgBox "Sorry. I don't know how to Undo. If you can show me how, please send me e-mail : kicheto@goatrance.com"
End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
End Sub

Private Sub mnuFileClose_Click()
   On Error Resume Next
   dbDataBase.Close
   mnuFileClose.Enabled = False
End Sub

Private Sub mnuFileOpen_Click()
Dim FileName As String
    On Error GoTo mnuFileOpenError:
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "MS Access DataBases (*.mdb)|*.mdb"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        FileName = .FileName
    End With
   
    frmShowRecordset.Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & FileName & ";"
       
    Set dbWorkSpace = DBEngine.Workspaces(0)
    Set dbDataBase = dbWorkSpace.OpenDatabase(FileName)
    'Now you can Create new table
    mnuInsertTable.Enabled = True
    '------------------ new Query
    mnuInsertQuery.Enabled = True
    '----------- delete tables or queries
    mnuEditDelete.Enabled = True
    '----------- rename tables or queries
    mnuEditRename.Enabled = True
    '----------- close the database
    mnuFileClose.Enabled = True
    'Input the tables from the Database in frmDBControl.List1
    Call InputTablesToListBox(frmDBControl.List1)
                     
    frmDBControl.Caption = "Database : " & dlgCommonDialog.FileTitle
    frmDBControl.Show
    
    Exit Sub

mnuFileOpenError:
MsgBox Err.Description, vbCritical, "QuarantineDB : Error Num." & Err.Number
End Sub

Private Sub mnuFileNew_Click()
Dim NewDBName As String
   On Error Resume Next
   dbWorkSpace.Close
   On Error GoTo mnuFileNewError
   'NewDBName = InputBox("DataBase Name : ", "QuarantineDB Create New DataBase")
   With dlgCommonDialog
        .DialogTitle = "Create Database"
        .CancelError = False
        .Filter = "MS Access DataBases (*.mdb)|*.mdb"
        .ShowSave
        If Len(.FileName) = 0 Then
           Exit Sub
        End If
        NewDBName = dlgCommonDialog.FileName
        Set dbWorkSpace = DBEngine.Workspaces(0)
        Set dbDataBase = dbWorkSpace.CreateDatabase(NewDBName, dbLangGeneral)
        
        Call InputTablesToListBox(frmDBControl.List1)
                     
        'Now you can Create new table
        mnuInsertTable.Enabled = True
        '------------------ new Query
        mnuInsertQuery.Enabled = True
        '----------- delete tables or queries
         mnuEditDelete.Enabled = True
        '----------- rename tables or queries
        mnuEditRename.Enabled = True
        '----------- close the database
        mnuFileClose.Enabled = True
        
        frmDBControl.Caption = "Database : " & dlgCommonDialog.FileTitle
        frmDBControl.Show
    End With
    Exit Sub
mnuFileNewError:
MsgBox Err.Description, vbCritical, "QuarantineDB Error Num." & Err.Number
End Sub

