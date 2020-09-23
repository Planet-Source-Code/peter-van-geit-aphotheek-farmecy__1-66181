VERSION 5.00
Begin VB.Form frmstart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "opstartscherm"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7755
   Begin VB.CommandButton CMDZOEK 
      Caption         =   "&Zoeken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      Picture         =   "frmstart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdsluit 
      Caption         =   "&Sluiten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      Picture         =   "frmstart.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox txtstraat 
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Text            =   "straat"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtvoornaam 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Text            =   "voornaam"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtnaam 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Text            =   "naam"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdklant 
      Caption         =   "&Klant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Picture         =   "frmstart.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox txtRIJKSREGISTERNR 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "rijksregisternr"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblstraat 
      Caption         =   "&straat"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblvoornaam 
      Caption         =   "&voornaam"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblnaam 
      Caption         =   "&naam"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblrijksregisternr 
      Caption         =   "&rijksregisternr"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private dbklant As DAO.Database
Private rsklant As DAO.Recordset
Private blnnewrec As Boolean
Private lngmaxrec As Long

Private Sub snewrec()
Dim dbnewrec As DAO.Database
Dim rsnewrec As DAO.Database
    Set dbnewrec = OpenDatabase(App.Path & "\klanten.mdb")
    Set rsnewrec = dbnewrec.OpenRecordset("tblarts", dbOpenTable)
    Call gsClearText(frm:=Me)
    With rsnewrec
    .MoveLast
    txtidnr.Text = .Fields("lngidnr") + 1
    .MoveFirst
    End With
    blnnewrec = True
    rsnewrec.Close
    dbnewrec.Close
    Set rsnewrec = Nothing
    Set dbnewrec = Nothing


End Sub

Public Sub Sleesrec(rslees As DAO.Recordset)
    Dim dbklant As DAO.Database
    Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)
 
    txtidnr.Text = rsklant.Fields("lngidnr").Value
    txtnaam.Text = rslees.Fields("strnaam").Value
    txtvoornaam.Text = rslees.Fields("strvoornaam").Value
    txtstraat.Text = rslees.Fields("strstraat").Value


End Sub
Private Function fnewid() As Long
    Dim rsnewid As DAO.Recordset
    Set rsnewid = dbadres.OpenRecordset("tblklant", dbOpenTable)
    With rsnewid
    .MoveLast
    fnewid = .Fields("lngidnr").Value + 1
    .Close
    End With
    Set rsnewid = Nothing
    
End Function

Private Sub cmdklant_Click()
frmklant.Show
End Sub

Private Sub cmdsluit_Click()
Unload Me
End Sub

Private Sub cmdzoek_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)

   With rsklant
   .Index = "primarykey"
   .Seek "=", Trim(txtrijksregisternr.Text)
   If Not .NoMatch Then
   txtnaam.Text = rsklant.Fields("strnaam").Value
   txtvoornaam.Text = rsklant.Fields("strvoornaam").Value
   txtstraat.Text = rsklant.Fields("strstraat").Value

       blnnewrec = False
   MsgBox "deze klant bestaat al"

       Else
   MsgBox "er is geen adres gevonden met id " & txtrijksregisternr.Text, vbOKOnly + vbInformation, "zoekresultaat"
   If MsgBox("wil je NIEUWE KLANT AANMAKEN", vbYesNo) = vbYes Then
   frmklant.Show
   End If



   .MoveLast
   Call gsClearText(frm:=Me)
   txtrijksregisternr.Text = .Fields("lngrijksregisternr").Value
   blnnewrec = True
   End If

   End With
   rsklant.Close
   dbklant.Close
   Set rsklant = Nothing
   Set dbklant = Nothing

    
    

End Sub


Private Sub Form_Load()
    Dim ctrl As Control
        For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Then
        ctrl.Text = vbNullString
    End If
    Next ctrl

Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset

blnnewrec = True

Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)
rsklant.MoveLast
lngmaxrec = rsklant.RecordCount
rsklant.MoveNext




End Sub

Private Sub txtidnr_KeyDown(KeyCode As Integer, Shift As Integer)
  
    Call gsClearText(frm:=Me)
End Sub

Private Sub txtidnr_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select

End Sub

Private Sub txtRIJKSREGISTERNR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select

End Sub
