VERSION 5.00
Begin VB.Form frmverkoop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verkoop"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txthoeveelheid9 
      Height          =   375
      Left            =   3480
      TabIndex        =   89
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Txthoeveelheid8 
      Height          =   375
      Left            =   3480
      TabIndex        =   88
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Txthoeveelheid7 
      Height          =   375
      Left            =   3480
      TabIndex        =   87
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Txthoeveelheid6 
      Height          =   375
      Left            =   3480
      TabIndex        =   86
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Txthoeveelheid5 
      Height          =   375
      Left            =   3480
      TabIndex        =   85
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Txthoeveelheid4 
      Height          =   375
      Left            =   3480
      TabIndex        =   84
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Txthoeveelheid3 
      Height          =   375
      Left            =   3480
      TabIndex        =   83
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Txthoeveelheid2 
      Height          =   375
      Left            =   3480
      TabIndex        =   82
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Txthoeveelheid1 
      Height          =   375
      Left            =   3480
      TabIndex        =   81
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Txtnaam9 
      Height          =   375
      Left            =   2040
      TabIndex        =   79
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Txtnaam8 
      Height          =   375
      Left            =   2040
      TabIndex        =   78
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Txtnaam7 
      Height          =   375
      Left            =   2040
      TabIndex        =   77
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Txtnaam6 
      Height          =   375
      Left            =   2040
      TabIndex        =   76
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtnaam5 
      Height          =   375
      Left            =   2040
      TabIndex        =   75
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Txtnaam4 
      Height          =   375
      Left            =   2040
      TabIndex        =   74
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Txtnaam3 
      Height          =   375
      Left            =   2040
      TabIndex        =   73
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Txtnaam2 
      Height          =   375
      Left            =   2040
      TabIndex        =   72
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Txtnaam1 
      Height          =   375
      Left            =   2040
      TabIndex        =   71
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox Cbo1 
      Height          =   315
      Left            =   120
      TabIndex        =   70
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Txtbetaald 
      Height          =   375
      Left            =   7200
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Cmdsubtotaal 
      Caption         =   "&Subtotaal"
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
      Left            =   5040
      Picture         =   "frmverkoop.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox txtaantal9 
      Height          =   405
      Left            =   6360
      TabIndex        =   55
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtaantal8 
      Height          =   405
      Left            =   6360
      TabIndex        =   54
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtaantal7 
      Height          =   405
      Left            =   6360
      TabIndex        =   53
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtaantal6 
      Height          =   375
      Left            =   6360
      TabIndex        =   52
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtaantal5 
      Height          =   375
      Left            =   6360
      TabIndex        =   51
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtaantal4 
      Height          =   375
      Left            =   6360
      TabIndex        =   50
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtaantal3 
      Height          =   375
      Left            =   6360
      TabIndex        =   49
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtaantal2 
      Height          =   375
      Left            =   6360
      TabIndex        =   48
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtaantal1 
      Height          =   375
      Left            =   6360
      TabIndex        =   47
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Txteenheidsprijs9 
      Height          =   375
      Left            =   4800
      TabIndex        =   46
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Txteenheidsprijs8 
      Height          =   375
      Left            =   4800
      TabIndex        =   45
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Txteenheidsprijs7 
      Height          =   375
      Left            =   4800
      TabIndex        =   44
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Txteenheidsprijs6 
      Height          =   375
      Left            =   4800
      TabIndex        =   43
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Txteenheidsprijs5 
      Height          =   375
      Left            =   4800
      TabIndex        =   42
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Txteenheidsprijs4 
      Height          =   375
      Left            =   4800
      TabIndex        =   41
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Txteenheidsprijs3 
      Height          =   375
      Left            =   4800
      TabIndex        =   40
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txteenheidsprijs2 
      Height          =   375
      Left            =   4800
      TabIndex        =   39
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txteenheidsprijs1 
      Height          =   375
      Left            =   4800
      TabIndex        =   38
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ComboBox cbo9 
      Height          =   315
      Left            =   120
      TabIndex        =   37
      Top             =   4920
      Width           =   1935
   End
   Begin VB.ComboBox cbo8 
      Height          =   315
      Left            =   120
      TabIndex        =   36
      Top             =   4560
      Width           =   1935
   End
   Begin VB.ComboBox cbo7 
      Height          =   315
      Left            =   120
      TabIndex        =   35
      Top             =   4200
      Width           =   1935
   End
   Begin VB.ComboBox cbo6 
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Top             =   3840
      Width           =   1935
   End
   Begin VB.ComboBox cbo5 
      Height          =   315
      Left            =   120
      TabIndex        =   33
      Top             =   3480
      Width           =   1935
   End
   Begin VB.ComboBox cbo4 
      Height          =   315
      Left            =   120
      TabIndex        =   32
      Top             =   3120
      Width           =   1935
   End
   Begin VB.ComboBox cbo3 
      Height          =   315
      Left            =   120
      TabIndex        =   31
      Top             =   2760
      Width           =   1935
   End
   Begin VB.ComboBox cbo2 
      Height          =   315
      Left            =   120
      TabIndex        =   30
      Top             =   2400
      Width           =   1935
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
      Left            =   9600
      Picture         =   "frmverkoop.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton cmdwissel 
      Caption         =   "Wisse&l"
      Height          =   735
      Left            =   3360
      Picture         =   "frmverkoop.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command19 
      Height          =   1215
      Left            =   10320
      Picture         =   "frmverkoop.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command18 
      Height          =   1215
      Left            =   10320
      Picture         =   "frmverkoop.frx":11F8
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      Height          =   1215
      Left            =   10320
      Picture         =   "frmverkoop.frx":163A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdtotaal 
      Caption         =   "&Totaal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      Picture         =   "frmverkoop.frx":1A7C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdmemo 
      Caption         =   "&Memo"
      Height          =   735
      Left            =   3360
      Picture         =   "frmverkoop.frx":20E6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdvoorschrift 
      Caption         =   "&Voorschrift"
      Height          =   735
      Left            =   1800
      Picture         =   "frmverkoop.frx":2528
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdaantal 
      Caption         =   "&Aantal"
      Height          =   735
      Left            =   240
      Picture         =   "frmverkoop.frx":296A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdbestel 
      Caption         =   "&Bestel"
      Height          =   735
      Left            =   1800
      Picture         =   "frmverkoop.frx":2A6C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdzoek 
      Caption         =   "&Zoek"
      Height          =   735
      Left            =   240
      Picture         =   "frmverkoop.frx":2EAE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Frame fraverkoop 
      Caption         =   "&Verkoop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8640
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      Begin VB.TextBox Txtverkoop1 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtverkoop 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command16 
         Height          =   495
         Left            =   2400
         Picture         =   "frmverkoop.frx":3000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdup 
         Height          =   495
         Left            =   2400
         Picture         =   "frmverkoop.frx":3442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraarts 
      Caption         =   "&Arts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.TextBox Txtartsid 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Text            =   "artsid"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtarts 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Text            =   "arts"
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame frapatient 
      Caption         =   "&Patient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtfederatie 
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Text            =   "federatie"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtkg1 
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         Text            =   "kg1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtkg2 
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Text            =   "kg2"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtrijksregisternr 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Text            =   "rijksregisternr"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtklantnaam 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Text            =   "klantnaam"
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Label Lblteruggave 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "wissel"
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
      Left            =   7200
      TabIndex        =   90
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Lblhoeveelheid 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "hoeveelheid"
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
      Left            =   3480
      TabIndex        =   80
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Lblbetaald 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&betaald"
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
      Left            =   7200
      TabIndex        =   69
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Lblwissel 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7200
      TabIndex        =   67
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Lbltotaal 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7200
      TabIndex        =   66
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Lblsubtotaal9 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   65
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Lblsubtotaal8 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   64
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Lblsubtotaal7 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   63
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Lblsubtotaal6 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   62
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Lblsubtotaal5 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   61
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Lblsubtotaal4 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   60
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Lblsubtotaal3 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   59
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label LBLsubtotaal2 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   58
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Lblsubtotaal1 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7200
      TabIndex        =   57
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblnaam 
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "naam"
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
      Left            =   2040
      TabIndex        =   29
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblsubtotaal 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subtotaal"
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
      Left            =   7200
      TabIndex        =   26
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblaantal 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aantal"
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
      Left            =   6360
      TabIndex        =   25
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lbleenheidsprijs 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Eenheidsprijs"
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
      Left            =   4800
      TabIndex        =   24
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblcnkcode 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "cnk code"
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
      TabIndex        =   23
      Top             =   1680
      Width           =   1935
   End
End
Attribute VB_Name = "frmverkoop"
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
    txtklantnaam.Text = rslees.Fields("strnaam").Value
    txtkg1.Text = rslees.Fields("strkg1").Value
    txtkg2.Text = rslees.Fields("strkg2").Value
    txtfederatie.Text = rslees.Fields("strfederatie").Value


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

Private Sub cmdsluiten_Click()


End Sub


Private Sub Cbo1_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
  With rsklant
    .Index = "primarykey"
    .Seek "=", Cbo1.List(Cbo1.ListIndex)
    
    If Not .NoMatch Then
      txteenheidsprijs1.Text = .Fields("eenheidsprijs").Value
      Txtnaam1.Text = .Fields("naam").Value
      Txthoeveelheid1.Text = .Fields("hoeveelheid").Value
      blnNieuwRec = False
          
      
    End If
  End With
  rsklant.Close
  dbklant.Close
  
  Set rsklant = Nothing
  Set dbklant = Nothing

End Sub

Private Sub cbo2_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
  With rsklant
    .Index = "primarykey"
    .Seek "=", cbo2.List(cbo2.ListIndex)
    
    If Not .NoMatch Then
      txteenheidsprijs2.Text = .Fields("eenheidsprijs").Value
      Txtnaam2.Text = .Fields("naam").Value
      Txthoeveelheid2.Text = .Fields("hoeveelheid").Value

      blnNieuwRec = False
          
      
    End If
  End With
  rsklant.Close
  dbklant.Close
  
  Set rsklant = Nothing
  Set dbklant = Nothing

End Sub

Private Sub cbo3_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
  With rsklant
    .Index = "primarykey"
    .Seek "=", cbo3.List(cbo3.ListIndex)
    
    If Not .NoMatch Then
      Txteenheidsprijs3.Text = .Fields("eenheidsprijs").Value
      Txtnaam3.Text = .Fields("naam").Value
      Txthoeveelheid3.Text = .Fields("hoeveelheid").Value

      blnNieuwRec = False
          
      
    End If
  End With
  rsklant.Close
  dbklant.Close
  
  Set rsklant = Nothing
  Set dbklant = Nothing

End Sub

Private Sub cbo4_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
  With rsklant
    .Index = "primarykey"
    .Seek "=", cbo4.List(cbo4.ListIndex)
    
    If Not .NoMatch Then
      Txteenheidsprijs4.Text = .Fields("eenheidsprijs").Value
      Txtnaam4.Text = .Fields("naam").Value
      Txthoeveelheid4.Text = .Fields("hoeveelheid").Value

      blnNieuwRec = False
          
      
    End If
  End With
  rsklant.Close
  dbklant.Close
  
  Set rsklant = Nothing
  Set dbklant = Nothing

End Sub

Private Sub cbo5_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
  With rsklant
    .Index = "primarykey"
    .Seek "=", cbo5.List(cbo5.ListIndex)
    
    If Not .NoMatch Then
      Txteenheidsprijs5.Text = .Fields("eenheidsprijs").Value
      txtnaam5.Text = .Fields("naam").Value
      Txthoeveelheid5.Text = .Fields("hoeveelheid").Value

      blnNieuwRec = False
          
      
    End If
  End With
  rsklant.Close
  dbklant.Close
  
  Set rsklant = Nothing
  Set dbklant = Nothing

End Sub

Private Sub cbo6_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
  With rsklant
    .Index = "primarykey"
    .Seek "=", cbo6.List(cbo6.ListIndex)
    
    If Not .NoMatch Then
      Txteenheidsprijs6.Text = .Fields("eenheidsprijs").Value
      Txtnaam6.Text = .Fields("naam").Value
      Txthoeveelheid6.Text = .Fields("hoeveelheid").Value

      blnNieuwRec = False
          
      
    End If
  End With
  rsklant.Close
  dbklant.Close
  
  Set rsklant = Nothing
  Set dbklant = Nothing

End Sub

Private Sub cbo7_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
  With rsklant
    .Index = "primarykey"
    .Seek "=", cbo7.List(cbo7.ListIndex)
    
    If Not .NoMatch Then
      Txteenheidsprijs7.Text = .Fields("eenheidsprijs").Value
      Txtnaam7.Text = .Fields("naam").Value
      Txthoeveelheid7.Text = .Fields("hoeveelheid").Value

      blnNieuwRec = False
          
      
    End If
  End With
  rsklant.Close
  dbklant.Close
  
  Set rsklant = Nothing
  Set dbklant = Nothing

End Sub

Private Sub cbo8_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
  With rsklant
    .Index = "primarykey"
    .Seek "=", cbo8.List(cbo8.ListIndex)
    
    If Not .NoMatch Then
      Txteenheidsprijs8.Text = .Fields("eenheidsprijs").Value
      Txtnaam8.Text = .Fields("naam").Value
      Txthoeveelheid8.Text = .Fields("hoeveelheid").Value

      blnNieuwRec = False
          
      
    End If
  End With
  rsklant.Close
  dbklant.Close
  
  Set rsklant = Nothing
  Set dbklant = Nothing

End Sub

Private Sub cbo9_Click()
Dim dbklant As DAO.Database
Dim rsklant As DAO.Recordset
   Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
   Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenTable)
  With rsklant
    .Index = "primarykey"
    .Seek "=", cbo9.List(cbo9.ListIndex)
    
    If Not .NoMatch Then
      Txteenheidsprijs9.Text = .Fields("eenheidsprijs").Value
      Txtnaam9.Text = .Fields("naam").Value
      Txthoeveelheid9.Text = .Fields("hoeveelheid").Value

      blnNieuwRec = False
          
      
    End If
  End With
  rsklant.Close
  dbklant.Close
  
  Set rsklant = Nothing
  Set dbklant = Nothing

End Sub

Private Sub cmdaantal_Click()
 txtaantal1.Text = vbNullString
 txtaantal2.Text = vbNullString
 txtaantal3.Text = vbNullString
 txtaantal4.Text = vbNullString
 txtaantal5.Text = vbNullString
 txtaantal6.Text = vbNullString
 txtaantal7.Text = vbNullString
 txtaantal8.Text = vbNullString
 txtaantal9.Text = vbNullString

End Sub

Private Sub cmdbestel_Click()
frmbestel.Show

End Sub


Private Sub cmdsluit_Click()
If MsgBox("opgelet is de rekening reeds BETAALD", vbYesNo) = vbYes Then

Unload Me
End If


End Sub

Private Sub Cmdsubtotaal_Click()
Lblsubtotaal1.Caption = Val(txteenheidsprijs1.Text) * Val(txtaantal1.Text)
LBLsubtotaal2.Caption = Val(txteenheidsprijs2.Text) * Val(txtaantal2.Text)
Lblsubtotaal3.Caption = Val(Txteenheidsprijs3.Text) * Val(txtaantal3.Text)
Lblsubtotaal4.Caption = Val(Txteenheidsprijs4.Text) * Val(txtaantal4.Text)
Lblsubtotaal5.Caption = Val(Txteenheidsprijs5.Text) * Val(txtaantal5.Text)
Lblsubtotaal6.Caption = Val(Txteenheidsprijs6.Text) * Val(txtaantal6.Text)
Lblsubtotaal7.Caption = Val(Txteenheidsprijs7.Text) * Val(txtaantal7.Text)
Lblsubtotaal8.Caption = Val(Txteenheidsprijs8.Text) * Val(txtaantal8.Text)
Lblsubtotaal9.Caption = Val(Txteenheidsprijs9.Text) * Val(txtaantal9.Text)

End Sub

Private Sub cmdtotaal_Click()
Lbltotaal.Caption = Val(Lblsubtotaal1.Caption) + Val(LBLsubtotaal2.Caption) + Val(Lblsubtotaal3.Caption) _
+ Val(Lblsubtotaal4.Caption) + Val(Lblsubtotaal5.Caption) + Val(Lblsubtotaal6.Caption) + Val(Lblsubtotaal7.Caption) _
+ Val(Lblsubtotaal8.Caption) + Val(Lblsubtotaal9.Caption)

End Sub

Private Sub cmdvoorschrift_Click()
frmvoorschrift.Show
End Sub

Private Sub cmdwissel_Click()
Lblwissel.Caption = Val(Txtbetaald.Text) - Val(Lbltotaal.Caption)

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
   txtklantnaam.Text = rsklant.Fields("strnaam").Value
    txtkg1.Text = rsklant.Fields("strkg1").Value
    txtkg2.Text = rsklant.Fields("strkg2").Value
    txtfederatie.Text = rsklant.Fields("strfederatie").Value
    txtarts.Text = rsklant.Fields("strarts").Value
    Txtartsid.Text = rsklant.Fields("artsid").Value
   Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)

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
Set rsklant = dbklant.OpenRecordset("tblproductnaam", dbOpenSnapshot)



If Not rsklant.BOF And Not rsklant.EOF Then

    rsklant.MoveFirst

    Do Until rsklant.EOF

    Cbo1.AddItem rsklant.Fields("cnkcode")
    cbo2.AddItem rsklant.Fields("cnkcode")
    cbo3.AddItem rsklant.Fields("cnkcode")
    cbo4.AddItem rsklant.Fields("cnkcode")
    cbo5.AddItem rsklant.Fields("cnkcode")
    cbo6.AddItem rsklant.Fields("cnkcode")
    cbo7.AddItem rsklant.Fields("cnkcode")
    cbo8.AddItem rsklant.Fields("cnkcode")
    cbo9.AddItem rsklant.Fields("cnkcode")

    rsklant.MoveNext
    
Loop

Else

End If

Set dbklant = OpenDatabase(App.Path & "\klanten.mdb")
Set rsklant = dbklant.OpenRecordset("tblklant", dbOpenTable)
rsklant.MoveLast
   Call gsClearText(frm:=Me)

lngmaxrec = rsklant.RecordCount
txtrijksregisternr.Text = rsklant.Fields("lngrijksregisternr").Value
rsklant.MoveFirst

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


Private Sub txtklantnaam_Validate(Cancel As Boolean)
If Len(Trim(ActiveControl)) > 0 Then
    ActiveControl.BackColor = vbWhite
    ActiveControl.ToolTipText = ""
Else
    ActiveControl.BackColor = vbRed
    ActiveControl.ToolTipText = "dit is een verplicht veld"
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


