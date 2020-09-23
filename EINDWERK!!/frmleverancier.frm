VERSION 5.00
Begin VB.Form frmleverancier 
   Caption         =   "Leverancier"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtapotheekid 
      Height          =   495
      Left            =   4560
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   960
      Width           =   3135
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
      Height          =   1215
      Left            =   9000
      Picture         =   "frmleverancier.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txttelefoon 
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Text            =   "Text7"
      Top             =   4920
      Width           =   3735
   End
   Begin VB.TextBox txtgemeente 
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtpostcode 
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   3600
      Width           =   3735
   End
   Begin VB.TextBox txtadres 
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox txtnaam 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Frame fracommand 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   8760
      TabIndex        =   2
      Top             =   360
      Width           =   2775
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
         Height          =   1215
         Left            =   360
         Picture         =   "frmleverancier.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2640
         Width           =   2055
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
         Height          =   1215
         Left            =   360
         Picture         =   "frmleverancier.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1440
         Width           =   2055
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
         Height          =   1215
         Left            =   360
         Picture         =   "frmleverancier.frx":0696
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox txtlevid 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label lblapotheekid 
      Caption         =   "&Apotheekid"
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
      Left            =   4560
      TabIndex        =   18
      Top             =   360
      Width           =   3135
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
      Left            =   360
      TabIndex        =   12
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblgemeente 
      Caption         =   "&Gemeente"
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
      Left            =   4680
      TabIndex        =   8
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label lblpostcode 
      Caption         =   "&postcode"
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
      TabIndex        =   7
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lbladres 
      Caption         =   "&adres"
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
      Left            =   4680
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblnaam 
      Caption         =   "&Naam"
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
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label lblleverancierid 
      Caption         =   "&leverancierID"
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
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmleverancier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private dbklant As DAO.Database
Private rsklant As DAO.Recordset
Private blnnewrec As Boolean
Private lngmaxrec As Long



Public Sub Sleesrec(rslees As DAO.Recordset)
    Dim dbklant As DAO.Database
    Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblleverancier", dbOpenTable)
 
    txtapotheekid.Text = rsklant.Fields("apotheekid").Value
   txtlevid.Text = rsklant.Fields("levid").Value
   txtnaam.Text = rsklant.Fields("naam").Value
   txtadres.Text = rsklant.Fields("adres").Value
   txtpostcode.Text = rsklant.Fields("postcode").Value
   txtgemeente.Text = rsklant.Fields("gemeente").Value
   txttelefoon.Text = rsklant.Fields("telefoon").Value


End Sub






Private Function fnewid() As Long
    Dim rsnewid As DAO.Recordset
    Set rsnewid = dbklant.OpenRecordset("tblleverancier", dbOpenTable)
    With rsnewid
    .MoveLast
    fnewid = .Fields("levid").Value
    .Close
    End With
    Set rsnewid = Nothing
    
End Function


Private Sub cmdbewaar_Click()
Dim blnonvolledigeinput As Boolean
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
Set rsklant = dbklant.OpenRecordset("tblleverancier", dbOpenTable)

If Len(Trim(txtlevid.Text)) = 0 Then
    txtapotheekid.BackColor = vbRed
    txtapotheekid.ToolTipText = "Het apotheekid is verplicht"
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
    rsklant.Fields("levid").Value = Trim(txtlevid.Text) & " "
    Else
    rsklant.Index = "primarykey"
    rsklant.Seek "=", Trim(txtlevid.Text)
    If Not rsklant.NoMatch Then
    rsklant.Edit
    Else
    Exit Sub
    End If
End If
     rsklant.Fields("apotheekid").Value = Trim(txtapotheekid.Text) & " "
     rsklant.Fields("levid").Value = Trim(txtlevid.Text) & " "
     rsklant.Fields("naam").Value = Trim(txtnaam.Text) & " "
     rsklant.Fields("adres").Value = Trim(txtadres.Text) & " "
     rsklant.Fields("postcode").Value = Trim(txtpostcode.Text) & " "
     rsklant.Fields("gemeente").Value = Trim(txtgemeente.Text) & " "
     rsklant.Fields("telefoon").Value = Trim(txttelefoon.Text) & " "
          MsgBox "input ok", vbOKOnly + vbInformation, "naam opslaan"


     rsklant.Update

     
     
    End If
    
    

End Sub


Private Sub cmdsluiten_Click()
If MsgBox("opgelet zijn de bestanden reeds BEWAART", vbYesNo) = vbYes Then

Unload Me
End If
End Sub



Private Sub cmdsluit_Click()
If MsgBox("opgelet zijn de bestanden reeds BEWAART", vbYesNo) = vbYes Then

Unload Me
End If
End Sub

Private Sub cmdwis_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
    Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
    Set rsklant = dbklant.OpenRecordset("tblleverancier", dbOpenTable)
    With rsklant
    .Index = "primarykey"
    .Seek "=", Trim(txtlevid.Text)
    If Not .NoMatch Then
        Call Sleesrec(rslees:=rsklant)
        If MsgBox("wil je record met ID " & txtlevid.Text, vbYesNo + vbQuestion + vbDefaultButton2, "verwijderen") = vbYes Then
        .Delete
        End If
    Else
    MsgBox "er is geen adres gevonden met id " & txtlevid.Text, vbOKOnly + vbInformation, "Zoekresultaat"
    End If
    .MoveLast
    Call gsClearText(frm:=Me)
    txtlevid.Text = .Fields("levid").Value
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
   Set rsklant = dbklant.OpenRecordset("tblleverancier", dbOpenTable)

   With rsklant
   .Index = "primarykey"
   .Seek "=", Trim(txtlevid.Text)
   If Not .NoMatch Then
   txtapotheekid.Text = rsklant.Fields("APOTHEEKID").Value
   txtnaam.Text = rsklant.Fields("naam").Value
   txtadres.Text = rsklant.Fields("adres").Value
   txtpostcode.Text = rsklant.Fields("postcode").Value
   txtgemeente.Text = rsklant.Fields("gemeente").Value
   txttelefoon.Text = rsklant.Fields("telefoon").Value
    
       blnnewrec = False
       Else
   MsgBox "er is geen adres gevonden met id " & txtlevid.Text, vbOKOnly + vbInformation, " zoekresultaat"
   .MoveLast
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
Set rsklant = dbklant.OpenRecordset("tblleverancier", dbOpenTable)
rsklant.MoveLast
   Call gsClearText(frm:=Me)

lngmaxrec = rsklant.RecordCount
txtlevid.Text = rsklant.Fields("levid").Value
rsklant.MoveNext



End Sub
Private Sub txtpostcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select

End Sub


Private Sub txtpostcode_Validate(Cancel As Boolean)
    If Len(Trim(ActiveControl)) > 0 Then
        ActiveControl.BackColor = vbWhite
        ActiveControl.ToolTipText = ""
    Else
        ActiveControl.BackColor = vbRed
        ActiveControl.ToolTipText = "Dit is een verplicht veld"
        Cancel = True
        End If

End Sub

