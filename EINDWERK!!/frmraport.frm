VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmrapport 
   Caption         =   "Rapporten"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19660801
      CurrentDate     =   37994
   End
   Begin VB.Frame fracommand 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   9000
      TabIndex        =   3
      Top             =   240
      Width           =   2535
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
         Left            =   240
         Picture         =   "frmraport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmraport.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton cmdafbeelden 
         Caption         =   "&Afbeelden"
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
         Picture         =   "frmraport.frx":07BC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9551
      _Version        =   393216
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
      Left            =   9120
      Picture         =   "frmraport.frx":0BFE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label lbldatum 
      Caption         =   "&Datum"
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
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmrapport"
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
    Set rsnewrec = dbnewrec.OpenRecordset("tblklant", dbOpenTable)
    Call gsClearText(frm:=Me)
    With rsnewrec
    .MoveLast
    txtrijksregisternr.Text = .Fields("lngRIJKSREGISTERNR") + 1
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
     rsklant.Fields("lngkaartnR").Value = Trim(txtkaartnr.Text) & " "
     rsklant.Fields("strarts").Value = Trim(txtarts.Text) & " "
          MsgBox "input ok", vbOKOnly + vbInformation, "naam opslaan"

     
     rsklant.Update
     
     
     Call snewrec
     
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



