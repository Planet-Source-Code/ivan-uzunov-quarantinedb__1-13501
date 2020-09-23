Attribute VB_Name = "Module1"
Public fMainForm As frmMain
Public adoRecordset As ADODB.Recordset
Public adoConnection As ADODB.Connection
'promenlivi neobhodimi za za suzdavane na obecti za dostup do bazi danni rabotqt s DAO
Public dbWorkSpace As Workspace
Public dbDataBase As Database
Public dbTableDef As TableDef
Public dbFieldNew As Field
Public dbTable As Recordset
Public dbQueryDef As QueryDef
Public dbQuery As Recordset

Public InputTheNewTable As Boolean 'dali nowatya tablica da se sloji v listbox
Public QueryName As String

Public Type FieldProperties
       FieldName As String
       Type As DatabaseTypeEnum
       Size As Byte
End Type

Sub Main()
    Dim fLogin As New frmLogin
    fLogin.Show vbModal
    If Not fLogin.OK Then
        'Login Failed so exit app
        End
    End If
    Unload fLogin


    Set fMainForm = New frmMain
    fMainForm.Show
End Sub

Public Sub InputTablesToListBox(List1 As ListBox)
'raboti s DAO
Dim TablesCount As Long
Dim TableName As String
Dim I As Integer
    On Error GoTo InputTablesToListBoxError
    List1.Clear
    'broqt na kolonite
    TablesCount = dbDataBase.TableDefs.Count
    'pupvite 6 ne sa za pokazvane (nekvi sturotii na access)
    For I = 0 To TablesCount - 1
        Set dbTableDef = dbDataBase.TableDefs(I)
        TableName = dbTableDef.Name
        'tova sa tablici na Access koito ne trqbva da se pipat
        If TableName <> "MSysAccessObjects" And TableName <> "MSysACEs" And TableName <> "MSysObjects" And TableName <> "MSysQueries" And TableName <> "MSysRelationships" Then
           List1.AddItem TableName
        End If
    Next I
    List1.Tag = "Tables"
    Exit Sub
InputTablesToListBoxError:
MsgBox Err.Description, vbCritical, "QuarantineDB : Error Num." & Err.Number
End Sub

Public Sub InputQueriesToListBox(List1 As ListBox)
'raboti s DAO
Dim QueriesCount As Long
Dim I As Integer
    On Error GoTo InputQueriesToListBoxError
    List1.Clear
    'broqt na kolonite
    QueriesCount = dbDataBase.QueryDefs.Count - 1
    'pupvite 6 ne sa za pokazvane (nekvi sturotii na access)
    For I = 0 To QueriesCount
        Set dbQueryDef = dbDataBase.QueryDefs(I)
        List1.AddItem dbQueryDef.Name
    Next I
    List1.Tag = "Queries"
    Exit Sub

InputQueriesToListBoxError:
MsgBox Err.Description, vbCritical, "QuarantineDB : Error Num." & Err.Number
End Sub

Public Function FindTypeConstant(strType As String) As Byte
    Select Case strType
           Case "Boolean": FindTypeConstant = 1
           Case "Byte": FindTypeConstant = 2
           Case "Integer": FindTypeConstant = 3
           Case "Long": FindTypeConstant = 4
           Case "Currency": FindTypeConstant = 5
           Case "Single": FindTypeConstant = 6
           Case "Double": FindTypeConstant = 7
           Case "Date/Time": FindTypeConstant = 8
           Case "Text": FindTypeConstant = 10
           Case "Binary": FindTypeConstant = 9
           Case "Memo": FindTypeConstant = 12
     End Select
End Function

Public Sub InputTablesToComboBox(ComboBox1 As ComboBox)
'raboti s DAO
Dim TablesCount As Long
Dim TableName As String
Dim I As Integer
    On Error GoTo InputTablesToComboBoxError
    ComboBox1.Clear
    'broqt na kolonite
    TablesCount = dbDataBase.TableDefs.Count
    'pupvite 6 ne sa za pokazvane (nekvi sturotii na access)
    For I = 0 To TablesCount - 1
        Set dbTableDef = dbDataBase.TableDefs(I)
        TableName = dbTableDef.Name
        'tova sa tablici na Access koito ne trqbva da se pipat
        If TableName <> "MSysAccessObjects" And TableName <> "MSysACEs" And TableName <> "MSysObjects" And TableName <> "MSysQueries" And TableName <> "MSysRelationships" Then
           ComboBox1.AddItem TableName
        End If
    Next I
    Exit Sub
InputTablesToComboBoxError:
MsgBox Err.Description, vbCritical, "QuarantineDB : Error Num." & Err.Number
End Sub
