VERSION 5.00
Begin VB.Form frmDBControl 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4890
   Begin VB.CommandButton cmdQueries 
      Caption         =   "Queries"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdTables 
      Caption         =   "Tables"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmDBControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub List1_Click()
'Dim FieldsCount As Integer
'Dim strSQL As String
'   With frmShowRecordset.Adodc1
'      strSQL = "Select "
'      FieldsCount = .Recordset.Fields.Count
'      For I = 0 To FieldsCount - 1
'          If I < FieldsCount - 1 Then
'             strSQL = strSQL & .Recordset.Fields(I).Name & ","
'          Else
'             strSQL = strSQL & .Recordset.Fields(I).Name
'          End If
'      Next I
'      strSQL = strSQL & " from " & List1.List(List1.ListIndex)
'      .CommandType = adCmdUnknown
'      .RecordSource = strSQL
'      .Refresh
'   End With
'   frmShowRecordset.Show
'End Sub

Private Sub cmdQueries_Click()
   Call InputQueriesToListBox(Me.List1)
End Sub

Private Sub cmdTables_Click()
   Call InputTablesToListBox(Me.List1)
End Sub

Private Sub Form_Load()
   'If you not do this the form is too big
   Me.Width = 5000
   Me.Height = 3000
   
   Call Form_Resize
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   List1.Move 960, 120, Me.ScaleWidth - 120, Me.ScaleHeight - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Close the database
   On Error Resume Next
   dbDataBase.Close
   'Now you can't create new table
   'frmMain.mnuInsertTable.Enabled = False
End Sub

Private Sub List1_DblClick()
Dim frmSRS As frmShowRecordset
   On Error GoTo List1StartRecordsetError
   If List1.ListIndex > -1 Then
       Me.MousePointer = vbHourglass
       'It doesn't matter if is dbTable or dbQuery
       Set dbTable = dbDataBase.OpenRecordset(List1.List(List1.ListIndex))
      
       Set frmSRS = New frmShowRecordset
              
       With frmSRS
           .Adodc1.CommandType = adCmdTable
           'Chanche the table
           .Adodc1.RecordSource = dbTable.Name
           'It must be refresh
           .Adodc1.Refresh
           .Caption = dbTable.Name
           .Show
       End With
       'We don't need dbTable any more
       dbTable.Close
       Me.MousePointer = vbDefault
    End If
    Exit Sub
List1StartRecordsetError:
MsgBox Err.Description, vbCritical, "QuarantineDB : Error Num." & Err.Number
Me.MousePointer = vbDefault
End Sub
