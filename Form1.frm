VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "DBFix 1.0"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin MSComDlg.CommonDialog CDL1 
         Left            =   4440
         Top             =   1920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   7215
      End
      Begin VB.CommandButton Command1 
         Caption         =   ".."
         Height          =   315
         Left            =   8400
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   495
         Left            =   5280
         TabIndex        =   5
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Compare Database with Schema"
         Height          =   495
         Left            =   5280
         TabIndex        =   4
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create Master Schema"
         Height          =   495
         Left            =   5280
         TabIndex        =   3
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   840
         Width           =   5055
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2760
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DATABASE"
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   870
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program is purpose to compare and fix two different database schema
' It use DBFixTmp.mdb to hold master schema
' so you just create Master-Schema and bring the execute and DBFixTmp.mdb files to your customer and
' compare it outthere without bringing your development database

Dim dbTemp As Database
Dim dbx As Database

Private Sub cmdCreate_Click()
   If UCase(Text1) = UCase(App.Path & "\DBFixTmp.mdb") Then
      MsgBox "You Cannot Make schema with Master-Schema file"
      Exit Sub
   End If
   If Text1 = "" Then
      MsgBox "Enter database name !", vbExclamation
      Exit Sub
   End If
   
   If MsgBox("You are going to override Master-Schema, are you sure ?", vbYesNo) = vbNo Then Exit Sub
   Text2.Text = ""
   CreateLocalTable
   
   If Not OpenAllDB() Then Exit Sub
   
   Dim rsloTb As Recordset
   Dim rsloQy As Recordset
   Dim rsloIx As Recordset
   
   Set rsloTb = dbTemp.OpenRecordset("select * from myTABLE")
   Set rsloQy = dbTemp.OpenRecordset("select * from myQUERY")
   Set rsloIx = dbTemp.OpenRecordset("select * from myINDEX")
   
   putText "Creating myTABLE, myINDEX"
   Dim i%, j%, k%
   For i = 0 To dbx.TableDefs.Count - 1
      If UCase(Mid(dbx.TableDefs(i).Name, 1, 2)) = "TB" Or UCase(Mid(dbx.TableDefs(i).Name, 1, 2)) = "LO" Then
         With dbx.TableDefs(i)
         For j = 0 To .Fields.Count - 1

            rsloTb.AddNew
            rsloTb("TableName") = .Name
            rsloTb("SeqNum") = .Fields(j).OrdinalPosition
            rsloTb("FieldName") = .Fields(j).Name
            rsloTb("FieldType") = .Fields(j).Type
            rsloTb("Attributes") = .Fields(j).Attributes
            rsloTb("Required") = .Fields(j).Required
            rsloTb("Size") = .Fields(j).Size
            rsloTb("AllowZeroLength") = .Fields(j).AllowZeroLength
            
            If .Fields(j).Type = 1 Or .Fields(j).Type = 4 Or .Fields(j).Type = 5 Then
               ' yes/no, number, currency
               rsloTb("DefaultValue") = 0
            ElseIf .Fields(j).Type = 10 Then
               ' Text
               rsloTb("DefaultValue") = .Fields(j).DefaultValue
            End If
            rsloTb.Update
         Next j
         For j = 0 To dbx.TableDefs(i).Indexes.Count - 1
            rsloIx.AddNew
            rsloIx("TableName") = .Name
            rsloIx("IndexName") = .Indexes(j).Name
            rsloIx("Fields") = .Indexes(j).Fields
            rsloIx("Primary") = .Indexes(j).Primary
            rsloIx("Unique") = .Indexes(j).Unique
            rsloIx.Update
         Next j
         End With
      End If
   Next i
   
   putText "Creating myQUERY"
   For i = 0 To dbx.QueryDefs.Count - 1
      rsloQy.AddNew
      rsloQy!QueryName = dbx.QueryDefs(i).Name
      rsloQy!QueryDef = dbx.QueryDefs(i).SQL
      rsloQy.Update
   Next

   putText "Finished !!"
   Exit Sub
Err:
   MsgBox Err.Description
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   CDL1.DefaultExt = "mdb"
   CDL1.InitDir = App.Path
   CDL1.Filter = "Microsoft Access (*.mdb)|*.mdb|All Files (*.*)|*.*"
   
   CDL1.ShowOpen
   If CDL1.FileName <> "" Then
      Text1.Text = CDL1.FileName
   End If
End Sub


Private Sub putText(teks As String)
   Text2.Text = Text2.Text & Chr(13) & Chr(10) & teks
End Sub

Private Function UpdNewStat(oldstat As String, At_Digit As Integer)
   UpdNewStat = Mid(oldstat, 1, At_Digit - 1) & "U" & Mid(oldstat, At_Digit + 1)
End Function

Private Sub CreateLocalTable()
   Dim NewTable As TableDef
   Dim NewField As Field
   
   ' if dbFixTmp already exist, delete it
   If Dir(App.Path & "\DBFixTmp.mdb") = "DBFixTmp.mdb" Then Kill App.Path & "\DBFixTmp.mdb"
   
   Set dbTemp = CreateDatabase(App.Path & "\DBFixTmp.mdb", dbLangGeneral)
   Set NewTable = Nothing
   Set NewTable = dbTemp.CreateTableDef("myTABLE")
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("TableName")
         NewField.Type = 10
         NewField.Attributes = 2
         NewField.Size = 50
         NewField.AllowZeroLength = True
         NewTable.Fields.Append NewField
         
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("SeqNum")
         NewField.Type = 4
         NewField.Attributes = 1
         NewField.Size = 4
         NewTable.Fields.Append NewField
         
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("FieldName")
         NewField.Type = 10
         NewField.Attributes = 2
         NewField.Size = 50
         NewField.AllowZeroLength = True
         NewTable.Fields.Append NewField
         
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("FieldType")
         NewField.Type = 4
         NewField.Attributes = 1
         NewField.Size = 4
         NewTable.Fields.Append NewField
         
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("Attributes")
         NewField.Type = 4
         NewField.Attributes = 1
         NewField.Size = 4
         NewTable.Fields.Append NewField
         
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("Required")
         NewField.Type = 1
         NewField.Attributes = 1
         NewField.Size = 1
         NewTable.Fields.Append NewField
   
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("Size")
         NewField.Type = 4
         NewField.Attributes = 1
         NewField.Size = 4
         NewTable.Fields.Append NewField
   
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("AllowZeroLength")
         NewField.Type = 1
         NewField.Attributes = 1
         NewField.Size = 1
         NewTable.Fields.Append NewField
   
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("DefaultValue")
         NewField.Type = 10
         NewField.Attributes = 2
         NewField.Size = 50
         NewField.AllowZeroLength = True
         NewTable.Fields.Append NewField
   
         Set NewField = Nothing
         Set NewField = NewTable.CreateField("Status")
         NewField.Type = 10
         NewField.Attributes = 2
         NewField.Size = 20
         NewField.AllowZeroLength = True
         NewTable.Fields.Append NewField
   
   dbTemp.TableDefs.Append NewTable

   Set NewTable = Nothing
   Set NewTable = dbTemp.CreateTableDef("myINDEX")
   NewTable.Fields.Append NewTable.CreateField("TableName", dbText, 50)
   NewTable.Fields.Append NewTable.CreateField("IndexName", dbText, 50)
   NewTable.Fields.Append NewTable.CreateField("Fields", dbText, 255)
   NewTable.Fields.Append NewTable.CreateField("Primary", dbBoolean, 1)
   NewTable.Fields.Append NewTable.CreateField("Unique", dbBoolean, 1)
   NewTable.Fields.Append NewTable.CreateField("Status", dbText, 20)
   dbTemp.TableDefs.Append NewTable

   Set NewTable = Nothing
   Set NewTable = dbTemp.CreateTableDef("myQUERY")
   NewTable.Fields.Append NewTable.CreateField("QueryName", dbText, 50)
   NewTable.Fields.Append NewTable.CreateField("QueryDef", dbMemo, 0)
   NewTable.Fields.Append NewTable.CreateField("Status", dbText, 20)
   dbTemp.TableDefs.Append NewTable
   dbTemp.Close
   
End Sub

Private Function OpenAllDB() As Boolean
   On Error Resume Next
   dbx.Close
   Set dbx = Nothing
   dbTemp.Close
   Set dbTemp = Nothing
   
   On Error GoTo Err
   Set dbx = OpenDatabase(Text1)
   Set dbTemp = OpenDatabase(App.Path & "\DBFixTmp.mdb")
   OpenAllDB = True
   
   Exit Function
Err:
   MsgBox Err.Number & Chr(13) & Err.Description
   OpenAllDB = False
End Function

Private Sub cmdCompare_Click()
   
   Text2.Text = ""
   If Text1 = "" Then
      MsgBox "Enter database name !", vbExclamation
      Exit Sub
   End If
   
   If Dir(App.Path & "\DBFixTmp.mdb") <> "DBFixTmp.mdb" Then
      MsgBox "Master schema Not found, must be in same directory as program"
      Exit Sub
   End If
   
   If UCase(Text1) = UCase(App.Path & "\DBFixTmp.mdb") Then
      MsgBox "Cannot compare with Master-Schema file"
      Exit Sub
   End If
   
   'open database
   If Not OpenAllDB() Then Exit Sub
   
   
   Screen.MousePointer = vbHourglass
   'Status set to NEW
   dbTemp.Execute "update myTABLE set [status] = 'NEW'"
   dbTemp.Execute "update myINDEX set [status] = 'NEW'"
   dbTemp.Execute "update myQUERY set [status] = 'NEW'"
   
   Dim rsloTb As Recordset
   Dim rsloQy As Recordset
   Dim rsloIx As Recordset
   Dim rsClone As Recordset
   Dim SQL As String
   
   Dim NewTable As TableDef
   Dim NewField As Field
   Dim NewQuery As QueryDef
   Dim NewIndex As Index
   Dim NewStat As String
   
   Dim isNewTable As Boolean
   Dim isNewField As Boolean
   Dim isNewIndex As Boolean
   Dim isNewQuery As Boolean
   
   Set rsloTb = dbTemp.OpenRecordset("select * from myTABLE")
   Set rsloQy = dbTemp.OpenRecordset("select * from myQUERY")
   Set rsloIx = dbTemp.OpenRecordset("select * from myINDEX")
   
   Dim i%, j%, k%, l%, mPB1%, mPB2%
   DoEvents
   putText "CHECKING ..."
   
   ' Checking index and TableDef
   For i = 0 To dbx.TableDefs.Count - 1
      PB1.Value = i / dbx.TableDefs.Count * 40
      DoEvents
      
      'Ignore system files
      If UCase(Mid(dbx.TableDefs(i).Name, 1, 4)) <> "MSYS" Then
         If rsloTb.RecordCount > 0 Then
            For j = 0 To dbx.TableDefs(i).Fields.Count - 1
               rsloTb.FindFirst _
               " [TableName] = '" & dbx.TableDefs(i).Name & "' and " & _
               " [FieldName] = '" & dbx.TableDefs(i).Fields(j).Name & "'"
               rsloTb.Edit
               If Not rsloTb.NoMatch Then
                  NewStat = "-----"
                  If rsloTb("FieldType") <> dbx.TableDefs(i).Fields(j).Type Then NewStat = UpdNewStat(NewStat, 1)
                  If rsloTb("Attributes") <> dbx.TableDefs(i).Fields(j).Attributes Then NewStat = UpdNewStat(NewStat, 2)
                 'If rsloTb("Required") <> dbx.TableDefs(i).Fields(j).Required Then NewStat = UpdNewStat(NewStat, 3)
                  If rsloTb("Size") <> dbx.TableDefs(i).Fields(j).Size Then NewStat = UpdNewStat(NewStat, 4)
                  If rsloTb("AllowZeroLength") <> dbx.TableDefs(i).Fields(j).AllowZeroLength Then NewStat = UpdNewStat(NewStat, 5)
                  rsloTb("Status") = NewStat
               End If
               rsloTb.Update
            Next j
         End If
         
         If rsloIx.RecordCount > 0 Then
            For j = 0 To dbx.TableDefs(i).Indexes.Count - 1
               rsloIx.FindFirst _
               " [TableName] = '" & dbx.TableDefs(i).Name & "' and " & _
               " [IndexName] = '" & dbx.TableDefs(i).Indexes(j).Name & "'"
               rsloIx.Edit
               If dbx.TableDefs(i).Indexes(j).Name = "Doc Id" Then
                  Debug.Print "Sa"
               End If
               If Not rsloIx.NoMatch Then
                  NewStat = "-----"
                  If UCase(rsloIx("Fields")) <> UCase(dbx.TableDefs(i).Indexes(j).Fields) Then NewStat = UpdNewStat(NewStat, 1)
                  If rsloIx("Primary") <> dbx.TableDefs(i).Indexes(j).Primary Then NewStat = UpdNewStat(NewStat, 2)
                  If rsloIx("Unique") <> dbx.TableDefs(i).Indexes(j).Unique Then NewStat = UpdNewStat(NewStat, 3)
                  rsloIx("Status") = NewStat
               End If
               rsloIx.Update
            Next j
         End If
      End If
   Next i
   
   ' 10%
   If rsloQy.RecordCount > 0 Then
      For i = 0 To dbx.QueryDefs.Count - 1
         PB1.Value = i / dbx.QueryDefs.Count * 10 + 40
         DoEvents
         rsloQy.FindFirst _
         " [QueryName] = '" & dbx.QueryDefs(i).Name & "'"
         rsloQy.Edit
         
         If Not rsloQy.NoMatch Then
            NewStat = "-----"
            If rsloQy!QueryDef <> dbx.QueryDefs(i).SQL Then NewStat = "U"
            rsloQy("Status") = NewStat
         End If
         rsloQy.Update
      Next
   End If
   
   '------------------------------------------------------------------
   ' Fix the database
   ' Progress Bar 30% untuk tabledef
   putText "COMPARING TABLEDEF ..."
   Set rsClone = rsloTb.Clone
   If rsloTb.RecordCount > 0 Then
      rsloTb.MoveFirst
      rsloTb.MoveLast
      mPB1 = rsloTb.RecordCount
      rsloTb.MoveFirst
      mPB2 = 0
      
      Do While Not rsloTb.EOF
         isNewTable = False
         mPB2 = mPB2 + 1
         PB1.Value = 50 + mPB2 / mPB1 * 40
         DoEvents
         If rsloTb("Status") = "NEW" Then
            ' New field or new Table
            isNewTable = True
            For i = 0 To dbx.TableDefs.Count - 1
               If dbx.TableDefs(i).Name = rsloTb!TableName Then isNewTable = False
            Next
            
            If isNewTable Then
               isNewTable = True
               Set NewTable = Nothing
               Set NewTable = dbx.CreateTableDef(rsloTb!TableName)
               putText "  Create Table :: " & rsloTb!TableName
            Else
               Set NewTable = Nothing
               Set NewTable = dbx.TableDefs(rsloTb!TableName)
            End If
            
            putText "  Create Field :: " & rsloTb!TableName & "." & rsloTb!FieldName
            
            Set NewField = Nothing
            Set NewField = NewTable.CreateField(rsloTb!FieldName)
            NewField.Type = rsloTb!FieldType
            NewField.Attributes = rsloTb!Attributes
            NewField.Size = rsloTb![Size]
            NewField.OrdinalPosition = rsloTb!SeqNum
            If rsloTb!FieldType = 10 Then NewField.AllowZeroLength = rsloTb![AllowZeroLength]
            NewTable.Fields.Append NewField
            
            If isNewTable Then dbx.TableDefs.Append NewTable
            ' Assign new field
            If rsloTb!FieldType = 1 Or rsloTb!FieldType = 4 Or rsloTb!FieldType = 5 Then
               SQL = "update " & rsloTb!TableName & " set [" & rsloTb!FieldName & "] = 0"
               dbx.Execute SQL
            ElseIf rsloTb!FieldType = 10 Then
               If IsNull(rsloTb!DefaultValue) Or rsloTb!DefaultValue = "" Then
                  SQL = "update " & rsloTb!TableName & " set [" & rsloTb!FieldName & "] = ''"
               Else
                  SQL = "update " & rsloTb!TableName & " set [" & rsloTb!FieldName & "] = " & rsloTb!DefaultValue & ""
               End If
               dbx.Execute SQL
            End If
            
   
         ElseIf rsloTb!Status <> "-----" Then
            putText "  Update Field :: " & rsloTb!TableName & "." & rsloTb!FieldName
            
            ' Update field
            ' Cannot do update field properties,
            ' so the tricks is rename oldfield, create new field
            ' and move the content from oldfield to newfield then delete oldfield
            
            Set NewTable = Nothing
            Set NewTable = dbx.TableDefs(rsloTb!TableName)
            
            ' Maybe indexes use that field, so delete indexe, later create again
            On Error Resume Next
            For l = 0 To dbx.TableDefs(rsloTb!TableName).Indexes.Count - 1
               NewTable.Indexes.Delete dbx.TableDefs(rsloTb!TableName).Indexes(l).Name
            Next l
   
            'rename
            Set NewField = Nothing
            Set NewField = NewTable.Fields(rsloTb!FieldName)
            NewField.Name = rsloTb!FieldName & "_x"
            
            'create new
            Set NewField = Nothing
            Set NewField = NewTable.CreateField(rsloTb!FieldName)
            NewField.Type = rsloTb!FieldType
            NewField.Attributes = rsloTb!Attributes
            NewField.Size = rsloTb![Size]
            If rsloTb!FieldType = 10 Then NewField.AllowZeroLength = rsloTb![AllowZeroLength]
            NewField.OrdinalPosition = rsloTb!SeqNum
            NewTable.Fields.Append NewField
            
            ' copying content
            SQL = "update " & rsloTb!TableName & " set [" & rsloTb!FieldName & "] = [" & rsloTb!FieldName & "_x]"
            dbx.Execute SQL
            
            ' delete old field
            NewTable.Fields.Delete rsloTb!FieldName & "_x"
            On Error GoTo Err
         End If
         rsloTb.Edit
         rsloTb("Status") = rsloTb("Status") & "  DONE"
         rsloTb.Update
         rsloTb.MoveNext
      Loop
   Else
      PB1.Value = 80
      DoEvents
   End If
   
   putText "COMPARING INDEXES ..."
   OpenAllDB
   
   Set rsloTb = dbTemp.OpenRecordset("select * from myTABLE")
   Set rsloQy = dbTemp.OpenRecordset("select * from myQUERY")
   Set rsloIx = dbTemp.OpenRecordset("select * from myINDEX")
   
   If rsloIx.RecordCount > 0 Then
      rsloIx.MoveFirst
      rsloIx.MoveLast
      mPB1 = rsloIx.RecordCount
      mPB2 = 0
      rsloIx.MoveFirst
      Do While Not rsloIx.EOF
         mPB2 = mPB2 + 1
         PB1.Value = 80 + mPB2 / mPB1 * 10
         DoEvents
         If rsloIx("Status") <> "-----" Then
            'New atau update
            'If update state just delete index and create new index
            Set NewTable = Nothing
            Set NewTable = dbx.TableDefs(rsloIx!TableName)
            
            If rsloIx("Status") <> "NEW" Then
               putText "  Update index :: " & rsloIx!TableName & "." & rsloIx!Indexname
               ' Delete old index, later create new index
               NewTable.Indexes.Delete rsloIx!Indexname
            Else
               putText "  Create index :: " & rsloIx!TableName & "." & rsloIx!Indexname
            End If
            
            Set NewIndex = Nothing
            Set NewIndex = NewTable.CreateIndex(rsloIx!Indexname)
            NewIndex.Fields = rsloIx("Fields")
            NewIndex.Primary = rsloIx("Primary")
            NewIndex.Unique = rsloIx("Unique")
            'set newindex
            dbx.TableDefs(rsloIx!TableName).Indexes.Append NewIndex
         
         End If
         rsloIx.MoveNext
      Loop
   Else
      PB1.Value = 90
      DoEvents
   End If
   
   putText "COMPARING QUERYDEF ..."
   ' Progress bar = 10%
   If rsloQy.RecordCount > 0 Then
      rsloQy.MoveFirst
      rsloQy.MoveLast
      mPB1 = rsloQy.RecordCount
      mPB2 = 0
      rsloQy.MoveFirst
      Do While Not rsloQy.EOF
         mPB2 = mPB2 + 1
         PB1.Value = 90 + mPB2 / mPB1 * 10
         DoEvents
         If rsloQy("Status") <> "-----" Then
            If rsloQy("Status") <> "NEW" Then
               ' If query is different, just delete it
               putText "  Update Query :: " & rsloQy!QueryName
               dbx.QueryDefs.Delete rsloQy!QueryName
            Else
               putText "  Create Query :: " & rsloQy!QueryName
            End If
            Set NewQuery = Nothing
            Set NewQuery = dbx.CreateQueryDef(rsloQy!QueryName, rsloQy!QueryDef)
         End If
         rsloQy.MoveNext
      Loop
   Else
      PB1.Value = 100
      DoEvents
   End If
   putText "DONE .."
   Screen.MousePointer = vbNormal
   
   Exit Sub
   
Err:
   MsgBox Err.Description
   Resume Next
End Sub

Private Sub Form_Load()
   Text1.Text = ""
   PB1.Value = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   dbx.Close
   Set dbx = Nothing
   
   dbTemp.Close
   Set dbTemp = Nothing
End Sub


