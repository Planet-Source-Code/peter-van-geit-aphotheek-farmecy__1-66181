VERSION 5.00
Begin VB.Form frmvoorschrift 
   Caption         =   "Voorschrift"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txtrijksregisternr 
      Height          =   495
      Left            =   3960
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   840
      Width           =   3255
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
      Picture         =   "frmvoorschritf.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Frame fracommand 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   8760
      TabIndex        =   10
      Top             =   240
      Width           =   2655
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
         Picture         =   "frmvoorschritf.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2760
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
         Height          =   1215
         Left            =   240
         Picture         =   "frmvoorschritf.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1560
         Width           =   2175
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
         Picture         =   "frmvoorschritf.frx":0696
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox txtomschrijving 
      Height          =   2415
      Left            =   240
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   4800
      Width           =   5655
   End
   Begin VB.TextBox txtdosering 
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox txtcnkcode 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox txtnaam 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtvoorschrift 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Lblrijksregisternr 
      Caption         =   "&Rijksregisternr"
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
      Left            =   3960
      TabIndex        =   15
      Top             =   240
      Width           =   3255
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
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label lbldosering 
      Caption         =   "&Dosering"
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
      TabIndex        =   6
      Top             =   2880
      Width           =   3015
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
      Left            =   4080
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
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
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblvoorschriftid 
      Caption         =   "&VoorschriftID"
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
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmvoorschrift"
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
   Set rsklant = dbklant.OpenRecordset("tblvoorschrift", dbOpenTable)
    txtvoorschriftid.Text = rslees.Fields("voorschriftid").Value
    txtRIJKSREGISTERNR.Text = rslees.Fields("lngrijksregisternr").Value
    txtnaam.Text = rslees.Fields("naam").Value
    txtcnkcode.Text = rslees.Fields("cnkcode").Value
    txtdosering.Text = rslees.Fields("dosering").Value
    txtomschrijving.Text = rslees.Fields("omschrijving").Value


End Sub






Private Function fnewid() As Long
    Dim rsnewid As DAO.Recordset
    Set rsnewid = dbadres.OpenRecordset("tblvoorschrift", dbOpenTable)
    With rsnewid
    .MoveLast
    fnewid = .Fields("lngRIJKSREGISTERNR").Value
    .Close
    End With
    Set rsnewid = Nothing
    
End Function


Private Sub cmdbewaar_Click()
Dim blnonvolledigeinput As Boolean
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
Set rsklant = dbklant.OpenRecordset("tblvoorschrift", dbOpenTable)

If Len(Trim(txtRIJKSREGISTERNR.Text)) = 0 Then
    txtRIJKSREGISTERNR.BackColor = vbRed
    txtRIJKSREGISTERNR.ToolTipText = "Het rijksregisternr is verplicht"
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
    rsklant.Fields("lngRIJKSREGISTERNR").Value = Trim(txtRIJKSREGISTERNR.Text) & " "
    Else
    rsklant.Index = "primarykey"
    rsklant.Seek "=", Trim(txtRIJKSREGISTERNR.Text)
    If Not rsklant.NoMatch Then
    rsklant.Edit
    Else
    Exit Sub
    End If
End If
     rsklant.Fields("voorschriftid").Value = Trim(txtvoorschrift.Text) & " "
     rsklant.Fields("lngrijksregisternr").Value = Trim(txtRIJKSREGISTERNR.Text) & " "
     rsklant.Fields("cnkcode").Value = Trim(txtcnkcode.Text) & " "
     rsklant.Fields("dosering").Value = Trim(txtdosering.Text) & " "
     rsklant.Fields("dosering").Value = Trim(txtdosering.Text) & " "
     rsklant.Fields("omschrijving").Value = Trim(txtomschrijving.Text) & " "
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
    Set rsklant = dbklant.OpenRecordset("tblvoorschrift", dbOpenTable)
    With rsklant
    .Index = "primarykey"
    .Seek "=", Trim(txtRIJKSREGISTERNR.Text)
    If Not .NoMatch Then
        Call Sleesrec(rslees:=rsklant)
        If MsgBox("wil je record met ID " & txtRIJKSREGISTERNR.Text, vbYesNo + vbQuestion + vbDefaultButton2, "verwijderen") = vbYes Then
        .Delete
        End If
    Else
    MsgBox "er is geen adres gevonden met id " & txtRIJKSREGISTERNR.Text, vbOKOnly + vbInformation, "Zoekresultaat"
    End If
    .MoveLast
    Call gsClearText(frm:=Me)
    txtRIJKSREGISTERNR.Text = .Fields("lngrijksregisternr").Value
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
   Set rsklant = dbklant.OpenRecordset("tblvoorschrift", dbOpenTable)

   With rsklant
   .Index = "primarykey"
   .Seek "=", Trim(txtRIJKSREGISTERNR.Text)
   If Not .NoMatch Then
    txtvoorschrift.Text = rsklant.Fields("voorschriftid").Value
    txtnaam.Text = rsklant.Fields("naam").Value
    txtcnkcode.Text = rsklant.Fields("cnkcode").Value
    txtdosering.Text = rsklant.Fields("dosering").Value
    txtomschrijving.Text = rsklant.Fields("omschrijving").Value

       blnnewrec = False
       Else
   MsgBox "er is geen adres gevonden met id " & txtRIJKSREGISTERNR.Text, vbOKOnly + vbInformation, " zoekresultaat"
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
Set rsklant = dbklant.OpenRecordset("tblvoorschrift", dbOpenTable)
rsklant.MoveLast
   Call gsClearText(frm:=Me)

lngmaxrec = rsklant.RecordCount
txtRIJKSREGISTERNR.Text = rsklant.Fields("lngrijksregisternr").Value
rsklant.MoveNext



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


