VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Conversor DBF a MDB"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Convertir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2820
      TabIndex        =   2
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3780
      TabIndex        =   1
      Top             =   180
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   3675
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   4020
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim StrTmp As String
    StrTmp = BrowseForFolder(Me.hWnd, "Indique la ruta de la base de datos DBF")
    If Trim(StrTmp) <> "" Then
        If Right(StrTmp, 1) <> "\" Then StrTmp = StrTmp & "\"
        Me.Text1.Text = StrTmp
    End If
End Sub

Private Sub Command3_Click()
Me.Command3.Enabled = False
Screen.MousePointer = vbHourglass
Dim DB1 As Database
Dim Db2 As Database
Dim ws As Workspace
Set ws = Workspaces(0)
Set DB1 = ws.OpenDatabase(Me.Text1.Text, True, True, "dBASE 5.0;")
If Dir(Me.Text1.Text & "access.mdb") <> "" Then
    Kill Me.Text1.Text & "access.mdb"
End If
Set Db2 = ws.CreateDatabase(Me.Text1.Text & "access.mdb", dbLangGeneral)
Dim Td As TableDef
Dim Rs2 As Recordset
Dim Rs1 As Recordset
Dim a As Long
Dim i As Integer
    For Each Td In DB1.TableDefs
        Me.Label1.Caption = Td.Name & "(Estructura)"
        Me.Label1.Refresh
        On Error Resume Next
        Err.Clear
        Set Rs1 = DB1.OpenRecordset(Td.Name)
        If Err.Number = 0 Then
            On Error GoTo 0
            Call CrearTabla(Rs1, Db2)
            'traspasar datos
            If Rs1.RecordCount > 0 Then
                a = 0
                Rs1.MoveLast
                Rs1.MoveFirst
                Set Rs2 = Db2.OpenRecordset("Select * from " & Rs1.Name & ";")
                Do Until Rs1.EOF = True
                    Me.Label1.Caption = Td.Name & "(Traspasando " & a & " de " & Rs1.RecordCount & ")"
                    Me.Label1.Refresh
                    Rs2.AddNew
                    For i = 0 To Rs1.Fields.Count - 1
                        Rs2.Fields(Rs1.Fields(i).Name).Value = Rs1.Fields(i).Value
                    Next i
                    Rs2.Update
                    Rs1.MoveNext
                    a = a + 1
                Loop
                Rs2.Close
            End If
            Rs1.Close
        Else
            Debug.Print "Error procesando tabla " & Td.Name
        End If
    Next Td
    Me.Label1.Caption = ""
    Me.Label1.Refresh
    Db2.Close
    DB1.Close
    MsgBox "Proceso concluido", vbInformation, "Aviso"
Screen.MousePointer = vbDefault
Me.Command3.Enabled = True

End Sub
Private Sub CrearTabla(Rst As Recordset, Db2 As Database)
    Dim Td As TableDef
    Dim Fd As Field
    Dim FdOrg As Field
    Set Td = Db2.CreateTableDef(Rst.Name)
        For Each FdOrg In Rst.Fields
            Set Fd = Td.CreateField(FdOrg.Name, FdOrg.Type, FdOrg.Size)
                Fd.Attributes = FdOrg.Attributes
                Fd.Required = FdOrg.Required
                Td.Fields.Append Fd
        Next FdOrg
    Db2.TableDefs.Append Td
End Sub
