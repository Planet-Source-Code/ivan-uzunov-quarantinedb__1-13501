VERSION 5.00
Begin VB.Form frmCreateTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Table"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuildTable 
      Caption         =   "Build Table"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemoveField 
      Caption         =   "Remove Field"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdADDField 
      Caption         =   "Add Field"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ListBox lstFields 
      Height          =   2400
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtFieldName 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtTableName 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Size"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "FieldName"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Table Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmCreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FieldProp() As FieldProperties

Private Sub cmbType_Click()
   Select Case cmbType.ListIndex
          Case 0: txtSize.Text = "1" 'Boolean
                          txtSize.Enabled = False
          Case 1: txtSize.Text = "1" 'Byte
                          txtSize.Enabled = False
          Case 2: txtSize.Text = "2" 'Integer
                          txtSize.Enabled = False
          Case 3: txtSize.Text = "4" 'Long
                          txtSize.Enabled = False
          Case 4: txtSize.Text = "8" '"Currency"
                          txtSize.Enabled = False
          Case 5: txtSize.Text = "4" '"Single"
                          txtSize.Enabled = False
          Case 6: txtSize.Text = "8" '"Double"
                          txtSize.Enabled = False
          Case 7: txtSize.Text = "8" '"Date/Time
                          txtSize.Enabled = False
          Case 8: txtSize.Text = "50" '"Text"
                          txtSize.Enabled = True
          Case 9: txtSize.Text = "0"  '"Binary"
                          txtSize.Enabled = False
          Case 10: txtSize.Text = "0" '"Memo"
                          txtSize.Enabled = False
    End Select
End Sub

Private Sub cmdADDField_Click()
Dim I As Integer
   On Error GoTo cmdADDFieldError
   If txtFieldName.Text = "" Then
      MsgBox "Inccorect Data", vbCritical, "Quarantine Error"
      txtFieldName.SetFocus
      Exit Sub
   End If
   
   'Check is the Field Size Type correct (only if Type is vbText)
   If Val(txtSize.Text) > 50 Or Val(txtSize.Text) < 1 Then
      MsgBox "Inccorect Data", vbCritical, "Quarantine Error"
      txtSize.SetFocus
      Exit Sub
   End If
   'Check is the field awready created
   For I = 0 To lstFields.ListCount - 1
       If lstFields.List(I) = txtFieldName.Text Then
          MsgBox "The Field: " & txtFieldName.Text & " - awready exist", vbInformation + vbOKOnly, "Quarantine"
          txtFieldName.SetFocus
          Exit Sub
       End If
   Next I
   'If is not add the new field
   lstFields.AddItem txtFieldName.Text
   'Add the field in the temporary array
    ReDim Preserve FieldProp(lstFields.ListCount)
        FieldProp(lstFields.ListCount).FieldName = txtFieldName.Text
        FieldProp(lstFields.ListCount).Size = Val(txtSize)
        FieldProp(lstFields.ListCount).Type = FindTypeConstant(cmbType.Text)
        
   'Prepare txtFieldName for next input
   txtFieldName.Text = ""
   txtFieldName.SetFocus
   
   Exit Sub
cmdADDFieldError:
MsgBox Err.Description, vbCritical, "QuarantineDB : Error Num." & Err.Number
End Sub

Private Sub cmdBuildTable_Click()
Dim I As Integer
    On Error GoTo cmdBuildTableError
    'proverka dali e vuvedeno ime za tablicata
    If txtTableName.Text = "" Then
       MsgBox "Invalid Table Name", vbCritical, "Quarantine"
       txtTableName.SetFocus
       Exit Sub
    End If
    'Create the table
    Set dbTableDef = dbDataBase.CreateTableDef(txtTableName.Text)
    'Add the fields to the table
    For I = 1 To lstFields.ListCount
        'Set dbFieldDef = dbTableDef.CreateField(lstFields.Text, cmbType.Text, txtSize.Text)
        Set dbFieldNew = dbTableDef.CreateField(FieldProp(I).FieldName, FieldProp(I).Type, FieldProp(I).Size)
        dbTableDef.Fields.Append dbFieldNew
    Next I
    'Add the table in the database
    dbDataBase.TableDefs.Append dbTableDef
    
    If InputTheNewTable = True Then
       InputTheNewTable = False
       Call InputTablesToListBox(frmDBControl.List1)
    End If
    
    Exit Sub
cmdBuildTableError:
MsgBox Err.Description, vbCritical, "QuarantineDB : Error Num." & Err.Number
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRemoveField_Click()
Dim I As Integer
   'premahva Item lstFields
   If lstFields.ListIndex > -1 Then
      FieldProp(lstFields.ListIndex + 1).FieldName = lstFields.List(lstFields.ListIndex)
      For I = lstFields.ListIndex + 1 To lstFields.ListCount - 1
          If I < lstFields.ListCount - 1 Then
             FieldProp(I) = FieldProp(I)
          End If
      Next I
    
      ReDim Preserve FieldProp(lstFields.ListCount - 1)
      
      lstFields.RemoveItem lstFields.ListIndex
   End If
End Sub

Private Sub Form_Load()
   'indecsite se izpolzvat za opredelqne na typa pri suzdavane na poleto
   't.e. vseki indeks otgovarq na suotvetniqt tip
   cmbType.AddItem "Boolean"
   cmbType.AddItem "Byte"
   cmbType.AddItem "Integer"
   cmbType.AddItem "Long"
   cmbType.AddItem "Currency"
   cmbType.AddItem "Single"
   cmbType.AddItem "Double"
   cmbType.AddItem "Date/Time"
   cmbType.AddItem "Text"
   cmbType.AddItem "Bynary"
   cmbType.AddItem "Memo"
   
   cmbType.ListIndex = 8 'Text
End Sub


