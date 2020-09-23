VERSION 5.00
Begin VB.Form frmCreateQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Query"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbCriteria 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCreateQuery 
      Caption         =   "Create Query"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtCriteria 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox cmbField 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox cmbTable 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Criteria"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Find"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Select Field"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Select Table"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmCreateQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbCriteria_Click()
   txtCriteria.Enabled = True
   txtCriteria.SetFocus
End Sub

Private Sub cmbField_Click()
   On Error Resume Next
   If cmbField.ListIndex > -1 Then
      cmbCriteria.Enabled = True
      'txtCriteria.SetFocus
   End If
End Sub

Private Sub cmbTable_Click()
Dim I As Integer
Dim FieldsCount As Long
   On Error GoTo cmbTableError
   If cmbTable.ListIndex > -1 Then
      Set dbTable = dbDataBase.OpenRecordset(cmbTable.List(cmbTable.ListIndex))
      FieldsCount = dbTable.Fields.Count - 1
      
      cmbField.Clear
      'Add the table fields in cmbField
      For I = 0 To FieldsCount
          cmbField.AddItem dbTable.Fields(I).Name
      Next I
      
      cmbField.Enabled = True
'      txtCriteria.DataFormat = dbTable.Fields(cmbTable.ListIndex).Type
   End If
   
   Exit Sub
cmbTableError:
MsgBox Err.Description, vbCritical, "QuarantineDB : Error Num." & Err.Number
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdCreateQuery_Click()
'I'm sorry but it is a bug hire
'I don't know how to separate the difrent formats
'This is the razon we can make only queries with text format
'If you have idea how to help me : kicheto@goatrance.com
Dim strSQL As String
   On Error GoTo cmdCreateQueryError
   Set dbQueryDef = dbDataBase.CreateQueryDef(QueryName)
   
   strSQL = "Select * From " & cmbTable.Text & " where " & cmbField.Text & cmbCriteria.Text & "'" & txtCriteria.Text & "'"
   dbQueryDef.SQL = strSQL
   Exit Sub
cmdCreateQueryError:
MsgBox Err.Description, vbCritical, "QuarantineDB : Error Num." & Err.Number
On Error Resume Next
dbDataBase.QueryDefs.Delete QueryName
End Sub

Private Sub Form_Load()
   Call InputTablesToComboBox(cmbTable)
   cmbCriteria.AddItem "="
   cmbCriteria.AddItem ">"
   cmbCriteria.AddItem ">="
   cmbCriteria.AddItem "<"
   cmbCriteria.AddItem "<="
   cmbCriteria.AddItem "<>"
End Sub

Private Sub txtCriteria_Change()
   If txtCriteria.Text <> "" Then
      cmdCreateQuery.Enabled = True
   Else
      cmdCreateQuery.Enabled = False
   End If
End Sub
