VERSION 5.00
Begin VB.Form frmarts 
   Caption         =   "Arts"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txttelefoon 
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Text            =   "Text5"
      Top             =   5640
      Width           =   4695
   End
   Begin VB.CommandButton cmdsluiten 
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
      Height          =   975
      Left            =   9840
      Picture         =   "frmarts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Frame fracommand 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   9720
      TabIndex        =   16
      Top             =   240
      Width           =   1935
      Begin VB.CommandButton cmdwis 
         Caption         =   "&Wissen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmarts.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdzoek 
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
         Height          =   975
         Left            =   120
         Picture         =   "frmarts.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdbewaar 
         Caption         =   "&Bewaren"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         Picture         =   "frmarts.frx":0986
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.TextBox txtemail 
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Text            =   "Text12"
      Top             =   5640
      Width           =   3495
   End
   Begin VB.TextBox txtgemeente 
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Text            =   "Text8"
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox txthuisnr 
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Text            =   "Text7"
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox txtnaam 
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox txtpostcode 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   4320
      Width           =   4695
   End
   Begin VB.TextBox txtstraat 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox txtvoornaam 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1920
      Width           =   4695
   End
   Begin VB.TextBox txtartsid 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label lbltelefoon 
      Caption         =   "&Telefoon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label lblemail 
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   14
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label lblgemeente 
      Caption         =   "Gemeente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label lblhuisnummer 
      Caption         =   "Huisnummer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Naam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblpostcode 
      Caption         =   "Postcode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblstraat 
      Caption         =   "Straat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblvoornaam 
      Caption         =   "Voornaam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblartsID 
      Caption         =   "ArtsID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmarts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dbklant As DAO.Database
Private rsklant As DAO.Recordset
Private blnnewrec As Boolean
Private lngmaxrec As Long

Private Function fnewid() As Long
    Dim rsnewid As DAO.Recordset
    Set rsnewid = dbklant.OpenRecordset("tblarts", dbOpenTable)
    With rsnewid
    .MoveLast
    fnewid = .Fields("artsid").Value + 1
    .Close
    End With
    Set rsnewid = Nothing
    
End Function


Public Sub Sleesrec(rslees As DAO.Recordset)
    Dim dbklant As DAO.Database
    Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblarts", dbOpenTable)
 
    txtartsid.Text = rslees.Fields("artsid").Value
    txtnaam.Text = rslees.Fields("Anaam").Value
    txtvoornaam.Text = rslees.Fields("Avoornaam").Value
    txtstraat.Text = rslees.Fields("Astraat").Value
    txthuisnr.Text = rslees.Fields("Ahuisnummer").Value
    txtpostcode.Text = rslees.Fields("Apostcode").Value
    txtgemeente.Text = rslees.Fields("Agemeente").Value
    txttelefoon.Text = rslees.Fields("Atelefoon").Value
    txtemail.Text = rslees.Fields("Aemail").Value


End Sub

Private Sub cmdbewaar_Click()
Dim blnonvolledigeinput As Boolean
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
Set rsklant = dbklant.OpenRecordset("tblarts", dbOpenTable)

If Len(Trim(txtartsid.Text)) = 0 Then
    txtartsid.BackColor = vbRed
    txtartsid.ToolTipText = "Het artsidnr is verplicht"
    blnonvolledigeinput = True
    End If
    
If Len(Trim(txtvoornaam.Text)) = 0 Then
    txtvoornaam.BackColor = vbRed
    txtvoornaam.ToolTipText = "Het veld voornaam is verplicht"
    blnonvolledigeinput = True
    End If

If Len(Trim(txtnaam.Text)) = 0 Then
    txtnaam.BackColor = vbRed
    txtnaam.ToolTipText = "Het veld naam is verplicht"
    blnonvolledigeinput = True
    End If
    
If blnonvolledigeinput Then
    MsgBox "Sorry, maar gelieve de invoer van de rood gekleurde velden" & vbCrLf & "aan te passen.", vbOKOnly + vbInformation, "Ingave fout"
    
    Else
    If blnnewrec Then
    rsklant.AddNew
    rsklant.Fields("artsid").Value = Trim(txtartsid.Text) & " "
    Else
    rsklant.Index = "primarykey"
    rsklant.Seek "=", Trim(txtartsid.Text)
    If Not rsklant.NoMatch Then
    rsklant.Edit
    Else
    Exit Sub
    End If
End If
     rsklant.Fields("artsid").Value = Trim(txtartsid.Text) & " "
     rsklant.Fields("Anaam").Value = Trim(txtnaam.Text) & " "
     rsklant.Fields("Avoornaam").Value = Trim(txtvoornaam.Text) & " "
     rsklant.Fields("Astraat").Value = Trim(txtstraat.Text) & " "
     rsklant.Fields("Ahuisnummer").Value = Trim(txthuisnr.Text) & " "
     rsklant.Fields("Apostcode").Value = Trim(txtpostcode.Text) & " "
     rsklant.Fields("Agemeente").Value = Trim(txtgemeente.Text) & " "
     rsklant.Fields("Atelefoon").Value = Trim(txttelefoon.Text) & " "
     rsklant.Fields("Aemail").Value = Trim(txtemail.Text) & " "
          MsgBox "input ok", vbOKOnly + vbInformation, "naam opslaan"

     
     rsklant.Update
     
     
     
    End If

End Sub

Private Sub cmdsluiten_Click()
If MsgBox("opgelet zijn de bestanden reeds BEWAART", vbYesNo) = vbYes Then

Unload Me
End If


End Sub

Private Sub cmdwis_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
    Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
    Set rsklant = dbklant.OpenRecordset("tblarts", dbOpenTable)
    With rsklant
    .Index = "primarykey"
    .Seek "=", Trim(txtartsid.Text)
    If Not .NoMatch Then
        Call Sleesrec(rslees:=rsklant)
        If MsgBox("wil je record met ID " & txtartsid, vbYesNo + vbQuestion + vbDefaultButton2, "SCHRAPPEN") = vbYes Then
        .Delete
        End If
    Else
    MsgBox "er is geen adres gevonden met id " & txtidnr.Text, vbOKOnly + vbInformation, "Zoekresultaat"
    End If
    .MoveLast
    Call gsClearText(frm:=Me)
    txtartsid.Text = .Fields("artsid").Value
    blnnewrec = True
    End With
    rsklant.Close
    dbklant.Close
    Set rsklant = Nothing
    Set dbklant = Nothing

End Sub

Private Sub cmdzoek_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
    Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
    Set rsklant = dbklant.OpenRecordset("tblarts", dbOpenTable)
    
    With rsklant
    .Index = "primarykey"
    .Seek "=", Trim(txtartsid.Text)
    If Not .NoMatch Then
        Call Sleesrec(rslees:=rsklant)
        blnnewrec = False
        Else
    MsgBox "er is geen adres gevonden met id &  txtidnr.Text, vbOKOnly + vbInformation, zoekresultaat"
    .MoveLast
    Call gsClearText(frm:=Me)
    txtartsid.Text = .Fields("artsid").Value + 1
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
Set rsklant = dbklant.OpenRecordset("tblarts", dbOpenTable)

rsklant.MoveLast
lngmaxrec = rsklant.RecordCount
txtartsid.Text = rsklant.Fields("artsid").Value
rsklant.MoveNext

Set rsklant = dbklant.OpenRecordset("tblafnemer", dbOpenTable)
Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)

End Sub

Private Sub txtartsid_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select

End Sub

Private Sub txtartsid_Validate(Cancel As Boolean)
    If Len(Trim(ActiveControl)) > 0 Then
        ActiveControl.BackColor = vbWhite
        ActiveControl.ToolTipText = ""
    Else
        ActiveControl.BackColor = vbRed
        ActiveControl.ToolTipText = "Dit is een verplicht veld"
        Cancel = True
        End If

End Sub

