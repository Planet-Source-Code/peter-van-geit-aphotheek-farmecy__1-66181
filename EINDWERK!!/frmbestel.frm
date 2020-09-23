VERSION 5.00
Begin VB.Form frmbestel 
   Caption         =   "bestelling"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txtbijwerkingen 
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   3240
      Width           =   7695
   End
   Begin VB.TextBox txtomschrijving 
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   1800
      Width           =   3495
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
      Left            =   9120
      Picture         =   "frmbestel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Frame fracommand 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   8760
      TabIndex        =   5
      Top             =   240
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
         Height          =   1095
         Left            =   360
         Picture         =   "frmbestel.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2520
         Width           =   2175
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
         Height          =   1095
         Left            =   360
         Picture         =   "frmbestel.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmdbestel 
         Caption         =   "&Bestel"
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
         Left            =   360
         Picture         =   "frmbestel.frx":0696
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox txtcnkcode 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtnaam 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label lblbijwerkingen 
      Caption         =   "bijwerkingen"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblcnk 
      Caption         =   "&CNK code"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
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
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "frmbestel"
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
    Set rsnewrec = dbnewrec.OpenRecordset("tblproductnaam", dbOpenTable)
    Call gsClearText(frm:=Me)
    With rsnewrec
    .MoveLast
    txtcnkcode.Text = .Fields("cnkcode") + 1
    .MoveNext
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
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
 
    txtcnkcode.Text = rslees.Fields("cnkcode").Value
    txtnaam.Text = rslees.Fields("naam").Value
    txtomschrijving.Text = rslees.Fields("omschrijving").Value
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
     rsklant.Fields("omschrijving").Value = Trim(txtomschrijving.Text) & " "
     rsklant.Fields("bijwerkingen").Value = Trim(txtbijwerkingen.Text) & " "
          MsgBox "input ok", vbOKOnly + vbInformation, "naam opslaan"

     
     rsklant.Update
     
     
     Call snewrec
     
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
    txtomschrijving.Text = rsklant.Fields("omschrijving").Value
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

