VERSION 5.00
Begin VB.Form frmproduct 
   Caption         =   "Product fiche"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtbijwerkingen 
      Height          =   1815
      Left            =   240
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5160
      Width           =   6375
   End
   Begin VB.TextBox txteenheidsprijs 
      Height          =   495
      Left            =   7080
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3840
      Width           =   2055
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
      Left            =   9480
      Picture         =   "frmproduct.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Frame fracommands 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   9240
      TabIndex        =   12
      Top             =   240
      Width           =   2535
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
         Left            =   240
         Picture         =   "frmproduct.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2880
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
         Left            =   240
         Picture         =   "frmproduct.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
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
         Left            =   240
         Picture         =   "frmproduct.frx":0696
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.TextBox txtaantal 
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txteenheid 
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txthoeveelheid 
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtnaam 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtomschrijving 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   3840
      Width           =   6375
   End
   Begin VB.TextBox txtcnkcode 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label lblbijwerking 
      Caption         =   "&Bijwerkingen"
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
      TabIndex        =   19
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lbleenheidsprijs 
      Caption         =   "&eenheidsprijs in €"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblaantal 
      Caption         =   "&aantal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lbleenheid 
      Caption         =   "&eenheid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblhoeveelheid 
      Caption         =   "&hoeveelheid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblnaam 
      Caption         =   "&naam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblomschrijving 
      Caption         =   "&Omschrijving"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblcnkcode 
      Caption         =   "&Cnk-code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmproduct"
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
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
 
    txtcnkcode.Text = rslees.Fields("cnkcode").Value
    txtnaam.Text = rslees.Fields("strnaam").Value
    txthoeveelheid.Text = rslees.Fields("hoeveelheid").Value
    txteenheid.Text = rslees.Fields("eenheid").Value
    txtaantal.Text = rslees.Fields("aantal").Value
    txtomschrijving.Text = rslees.Fields("omschrijving").Value
    txteenheidsprijs.Text = rslees.Fields("eenheidsprijs").Value
    txtbijwerkingen.Text = rslees.Fields("bijwerkingen").Value


End Sub






Private Function fnewid() As Long
    Dim rsnewid As DAO.Recordset
    Set rsnewid = dbadres.OpenRecordset("tblproductnaam", dbOpenTable)
    With rsnewid
    .MoveLast
    fnewid = .Fields("cnkcode").Value
    .Close
    End With
    Set rsnewid = Nothing
    
End Function


Private Sub cmdbewaar_Click()
Dim blnonvolledigeinput As Boolean
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)

If Len(Trim(txtcnkcode.Text)) = 0 Then
    txtcnkcode.BackColor = vbRed
    txtcnkcode.ToolTipText = "de cnkcode is verplicht"
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
    rsklant.Fields("cnkcode").Value = Trim(txtcnkcode.Text) & " "
    Else
    rsklant.Index = "primarykey"
    rsklant.Seek "=", Trim(txtcnkcode.Text)
    If Not rsklant.NoMatch Then
    rsklant.Edit
    Else
    Exit Sub
    End If
End If
     rsklant.Fields("cnkcode").Value = Trim(txtcnkcode.Text) & " "
     rsklant.Fields("naam").Value = Trim(txtnaam.Text) & " "
     rsklant.Fields("hoeveelheid").Value = Trim(txthoeveelheid.Text) & " "
     rsklant.Fields("eenheid").Value = Trim(txteenheid.Text) & " "
     rsklant.Fields("aantal").Value = Trim(txtaantal.Text) & " "
     rsklant.Fields("omschrijving").Value = Trim(txtomschrijving.Text) & " "
     rsklant.Fields("eenheidsprijs").Value = Trim(txteenheidsprijs.Text) & " "
     rsklant.Fields("bijwerkingen").Value = Trim(txtbijwerkingen.Text) & " "
          MsgBox "input ok", vbOKOnly + vbInformation, "product opslaan"

     
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
    Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
    With rsklant
    .Index = "primarykey"
    .Seek "=", Trim(txtcnkcode.Text)
    If Not .NoMatch Then
        Call Sleesrec(rslees:=rsklant)
        If MsgBox("wil je record met ID " & txtcnkcode.Text, vbYesNo + vbQuestion + vbDefaultButton2, "verwijderen") = vbYes Then
        .Delete
        End If
    Else
    MsgBox "er is geen adres gevonden met id " & txtcnkcode.Text, vbOKOnly + vbInformation, "Zoekresultaat"
    End If
    .MoveLast
    Call gsClearText(frm:=Me)
    txtcnkcode.Text = .Fields("cnkcode").Value
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
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)

   With rsklant
   .Index = "primarykey"
   .Seek "=", Trim(txtcnkcode.Text)
   If Not .NoMatch Then
   txtnaam.Text = rsklant.Fields("naam").Value
   txthoeveelheid.Text = rsklant.Fields("hoeveelheid").Value
   txteenheid.Text = rsklant.Fields("eenheid").Value
   txtaantal.Text = rsklant.Fields("aantal").Value
   txtomschrijving.Text = rsklant.Fields("omschrijving").Value
   txteenheidsprijs.Text = rsklant.Fields("eenheidsprijs").Value
   txtbijwerkingen.Text = rsklant.Fields("bijwerkingen").Value

       blnnewrec = False
       Else
   MsgBox "er is geen adres gevonden met id " & txtcnkcode.Text, vbOKOnly + vbInformation, " zoekresultaat"
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
Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
rsklant.MoveLast
   Call gsClearText(frm:=Me)

lngmaxrec = rsklant.RecordCount
txtcnkcode.Text = rsklant.Fields("cnkcode").Value
rsklant.MoveNext



End Sub

Private Sub txtcnkcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select

End Sub


Private Sub txtnaam_Validate(Cancel As Boolean)
If Len(Trim(ActiveControl)) > 0 Then
    ActiveControl.BackColor = vbWhite
    ActiveControl.ToolTipText = ""
Else
    ActiveControl.BackColor = vbRed
    ActiveControl.ToolTipText = "dit is een verplicht veld"
    Cancel = True
End If


End Sub

