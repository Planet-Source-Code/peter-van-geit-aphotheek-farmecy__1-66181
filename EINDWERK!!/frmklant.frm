VERSION 5.00
Begin VB.Form frmklant 
   Caption         =   "klant"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
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
      Picture         =   "frmklant.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Frame fracommand 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   9720
      TabIndex        =   26
      Top             =   480
      Width           =   1935
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
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
         Left            =   120
         Picture         =   "frmklant.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2880
         Width           =   1695
      End
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
         Picture         =   "frmklant.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "frmklant.frx":0BAE
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Picture         =   "frmklant.frx":0FF0
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.TextBox txtfederatie 
      Height          =   375
      Left            =   8400
      TabIndex        =   21
      Text            =   "Text13"
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtarts 
      Height          =   375
      Left            =   5760
      TabIndex        =   25
      Text            =   "Text12"
      Top             =   6720
      Width           =   3495
   End
   Begin VB.TextBox txtkaartnr 
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Text            =   "Text11"
      Top             =   6720
      Width           =   4695
   End
   Begin VB.TextBox txtkg2 
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      Text            =   "Text10"
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtkg1 
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Text            =   "Text9"
      Top             =   5520
      Width           =   855
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
      Top             =   3240
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
   Begin VB.TextBox txtrizivnr 
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Text            =   "Text5"
      Top             =   5640
      Width           =   4695
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
   Begin VB.TextBox txtrijksregisternr 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label lblfederatie 
      Caption         =   "Federatie"
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
      Left            =   8280
      TabIndex        =   17
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lblarts 
      Caption         =   "Arts"
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
      TabIndex        =   23
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label lblkaartnr 
      Caption         =   "Kaartnummer"
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
      TabIndex        =   22
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label lblkg2 
      Caption         =   "KG2"
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
      Left            =   7080
      TabIndex        =   16
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblkg1 
      Caption         =   "KG1"
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
      TabIndex        =   15
      Top             =   5040
      Width           =   855
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
   Begin VB.Label lblnaam 
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
   Begin VB.Label lblrizivnr 
      Caption         =   "Rizivnr"
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
      TabIndex        =   14
      Top             =   5040
      Width           =   1815
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
   Begin VB.Label lblrijksregister 
      Caption         =   "Rijksregister nr"
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
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmklant"
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
   Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)
 
    txtrijksregisternr.Text = rslees.Fields("lngrijksregisternr").Value
    txtnaam.Text = rslees.Fields("strnaam").Value
    txtvoornaam.Text = rslees.Fields("strvoornaam").Value
    txtstraat.Text = rslees.Fields("strstraat").Value
    txthuisnr.Text = rslees.Fields("strhuisnummer").Value
    txtpostcode.Text = rslees.Fields("strpostcode").Value
    txtgemeente.Text = rslees.Fields("strgemeente").Value
    txtrizivnr.Text = rslees.Fields("lngrizivnr").Value
    txtkg1.Text = rslees.Fields("strkg1").Value
    txtkg2.Text = rslees.Fields("strkg2").Value
    txtfederatie.Text = rslees.Fields("strfederatie").Value
    txtkaartnr.Text = rslees.Fields("lngkaartnr").Value
    txtarts.Text = rslees.Fields("strarts").Value


End Sub






Private Function fnewid() As Long
    Dim rsnewid As DAO.Recordset
    Set rsnewid = dbadres.OpenRecordset("tblklant", dbOpenTable)
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
Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)

If Len(Trim(txtrijksregisternr.Text)) = 0 Then
    txtrijksregisternr.BackColor = vbRed
    txtrijksregisternr.ToolTipText = "Het rijksregisternr is verplicht"
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
    rsklant.Fields("lngRIJKSREGISTERNR").Value = Trim(txtrijksregisternr.Text) & " "
    Else
    rsklant.Index = "primarykey"
    rsklant.Seek "=", Trim(txtrijksregisternr.Text)
    If Not rsklant.NoMatch Then
    rsklant.Edit
    Else
    Exit Sub
    End If
End If
     rsklant.Fields("lngrijksregisternr").Value = Trim(txtrijksregisternr.Text) & " "
     rsklant.Fields("strnaam").Value = Trim(txtnaam.Text) & " "
     rsklant.Fields("strvoornaam").Value = Trim(txtvoornaam.Text) & " "
     rsklant.Fields("strstraat").Value = Trim(txtstraat.Text) & " "
     rsklant.Fields("strhuisnummer").Value = Trim(txthuisnr.Text) & " "
     rsklant.Fields("strpostcode").Value = Trim(txtpostcode.Text) & " "
     rsklant.Fields("strgemeente").Value = Trim(txtgemeente.Text) & " "
     rsklant.Fields("lngrizivnr").Value = Trim(txtrizivnr.Text) & " "
     rsklant.Fields("strkg1").Value = Trim(txtkg1.Text) & " "
     rsklant.Fields("strkg2").Value = Trim(txtkg2.Text) & " "
     rsklant.Fields("strfederatie").Value = Trim(txtfederatie.Text) & " "
     rsklant.Fields("lngkaartnr").Value = Trim(txtkaartnr.Text) & " "
     rsklant.Fields("strarts").Value = Trim(txtarts.Text) & " "
          MsgBox "input ok", vbOKOnly + vbInformation, "naam opslaan"

     
     rsklant.Update
     
     
     
    End If
    
    

End Sub

Private Sub cmdprint_Click()
Dim strhuidigefont As String, intfontsize As Integer
With Printer
strhuidigefont = .FontName
intfontsize = .FontSize
.ScaleMode = vbMillimeters
.CurrentX = 50
.CurrentY = 20
.FontName = "tahoma"
.FontSize = 32
.FontBold = Not .FontBold
Printer.Print "KLANT INFORMATIE"
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 40
Printer.Print lblrijksregister;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 40
Printer.Print txtrijksregisternr;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 45
Printer.Print lblvoornaam;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 45
Printer.Print txtvoornaam;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 50
Printer.Print lblnaam;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 50
Printer.Print txtnaam;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 55
Printer.Print lblstraat;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 55
Printer.Print txtstraat;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 60
Printer.Print lblhuisnummer;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 60
Printer.Print txthuisnr;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 65
Printer.Print lblpostcode;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 65
Printer.Print txtpostcode;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 70
Printer.Print lblgemeente;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 70
Printer.Print txtgemeente;

.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 75
Printer.Print lblrizivnr;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 75
Printer.Print txtrizivnr;

.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 80
Printer.Print lblkg1;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 80
Printer.Print txtkg1;

.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 85
Printer.Print lblkg2;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 85
Printer.Print txtkg2;

.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 90
Printer.Print lblfederatie;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 90
Printer.Print txtfederatie;

.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 95
Printer.Print lblkaartnr;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 95
Printer.Print txtkaartnr;

.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 25
.CurrentY = 100
Printer.Print lblarts;
.FontName = strhuidigefont
.FontSize = intfontsize
.CurrentX = 75
.CurrentY = 100
Printer.Print txtarts;


.EndDoc
End With

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
    Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)
    With rsklant
    .Index = "primarykey"
    .Seek "=", Trim(txtrijksregisternr.Text)
    If Not .NoMatch Then
        Call Sleesrec(rslees:=rsklant)
        If MsgBox("wil je record met ID " & txtrijksregisternr.Text, vbYesNo + vbQuestion + vbDefaultButton2, "verwijderen") = vbYes Then
        .Delete
        End If
    Else
    MsgBox "er is geen adres gevonden met id " & txtrijksregisternr.Text, vbOKOnly + vbInformation, "Zoekresultaat"
    End If
    .MoveLast
    Call gsClearText(frm:=Me)
    txtrijksregisternr.Text = .Fields("lngrijksregisternr").Value
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
   Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)

   With rsklant
   .Index = "primarykey"
   .Seek "=", Trim(txtrijksregisternr.Text)
   If Not .NoMatch Then
   txtnaam.Text = rsklant.Fields("strnaam").Value
   txtvoornaam.Text = rsklant.Fields("strvoornaam").Value
   txtstraat.Text = rsklant.Fields("strstraat").Value
   txthuisnr.Text = rsklant.Fields("strhuisnummer").Value
   txtpostcode.Text = rsklant.Fields("strpostcode").Value
   txtgemeente.Text = rsklant.Fields("strgemeente").Value
   txtrizivnr.Text = rsklant.Fields("lngrizivnr").Value
   txtkg1.Text = rsklant.Fields("strkg1").Value
   txtkg2.Text = rsklant.Fields("strkg2").Value
   txtfederatie.Text = rsklant.Fields("strfederatie").Value
   txtkaartnr.Text = rsklant.Fields("lngkaartnr").Value
   txtarts.Text = rsklant.Fields("strarts").Value

       blnnewrec = False
       Else
   MsgBox "er is geen adres gevonden met id " & txtrijksregisternr.Text, vbOKOnly + vbInformation, " zoekresultaat"
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
Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)
rsklant.MoveLast
   Call gsClearText(frm:=Me)

lngmaxrec = rsklant.RecordCount
txtrijksregisternr.Text = rsklant.Fields("lngrijksregisternr").Value
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


Private Sub txtarts_Validate(Cancel As Boolean)
If Len(Trim(ActiveControl)) > 0 Then
    
    ActiveControl.BackColor = vbWhite
    ActiveControl.ToolTipText = ""
Else
    ActiveControl.BackColor = vbRed
    ActiveControl.ToolTipText = "dit is een verplicht veld"
    Cancel = True
End If


End Sub


Private Sub txtkaartnr_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select

End Sub


Private Sub txtkaartnr_Validate(Cancel As Boolean)
    If Len(Trim(ActiveControl)) > 0 Then
        ActiveControl.BackColor = vbWhite
        ActiveControl.ToolTipText = ""
    Else
        ActiveControl.BackColor = vbRed
        ActiveControl.ToolTipText = "Dit is een verplicht veld"
        Cancel = True
        End If

End Sub


Private Sub txtkg1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select

End Sub


Private Sub txtkg1_Validate(Cancel As Boolean)
    If Len(Trim(ActiveControl)) > 0 Then
        ActiveControl.BackColor = vbWhite
        ActiveControl.ToolTipText = ""
    Else
        ActiveControl.BackColor = vbRed
        ActiveControl.ToolTipText = "Dit is een verplicht veld"
        Cancel = True
        End If

End Sub


Private Sub txtkg2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select

End Sub


Private Sub txtkg2_Validate(Cancel As Boolean)
    If Len(Trim(ActiveControl)) > 0 Then
        ActiveControl.BackColor = vbWhite
        ActiveControl.ToolTipText = ""
    Else
        ActiveControl.BackColor = vbRed
        ActiveControl.ToolTipText = "Dit is een verplicht veld"
        Cancel = True
        End If

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


Private Sub txtRIJKSREGISTERNR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select
    
            
            
            
End Sub


Private Sub txtrijksregisternr_Validate(Cancel As Boolean)
    If Len(Trim(ActiveControl)) > 0 Then
        ActiveControl.BackColor = vbWhite
        ActiveControl.ToolTipText = ""
    Else
        ActiveControl.BackColor = vbRed
        ActiveControl.ToolTipText = "Dit is een verplicht veld"
        Cancel = True
        End If
        
        
End Sub


Private Sub txtrizivnr_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
        KeyAscii = 0
        
        MsgBox "Sorry alleen getallen zijn geldig", vbOKOnly + vbInformation, "Foutieve ingave"
    End Select

End Sub


Private Sub txtrizivnr_Validate(Cancel As Boolean)
    If Len(Trim(ActiveControl)) > 0 Then
        ActiveControl.BackColor = vbWhite
        ActiveControl.ToolTipText = ""
    Else
        ActiveControl.BackColor = vbRed
        ActiveControl.ToolTipText = "Dit is een verplicht veld"
        Cancel = True
        End If

End Sub


Private Sub txtvoornaam_Validate(Cancel As Boolean)
If Len(Trim(ActiveControl)) > 0 Then
    ActiveControl.BackColor = vbWhite
    ActiveControl.ToolTipText = ""
Else
    ActiveControl.BackColor = vbRed
    ActiveControl.ToolTipText = "dit is een verplicht veld"
    Cancel = True
End If


End Sub


