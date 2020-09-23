VERSION 5.00
Begin VB.MDIForm mdiapotheek 
   BackColor       =   &H8000000C&
   Caption         =   "@potheek"
   ClientHeight    =   3975
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6915
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Menu mnubestand 
      Caption         =   "&Bestand"
      Begin VB.Menu mnuopen 
         Caption         =   "&Openen"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Sluiten"
      End
      Begin VB.Menu mnuopslaan 
         Caption         =   "O&pslaan"
      End
      Begin VB.Menu mnuafdrukken 
         Caption         =   "&Afdrukken"
      End
   End
   Begin VB.Menu mnuformulieren 
      Caption         =   "&Formulieren"
      Begin VB.Menu mnustart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuarts 
         Caption         =   "&Arts"
      End
      Begin VB.Menu mnubestel 
         Caption         =   "&Bestel"
      End
      Begin VB.Menu mnufactuur 
         Caption         =   "&Factuur"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuklant 
         Caption         =   "&Klant"
      End
      Begin VB.Menu mnuleverancier 
         Caption         =   "&Leverancier"
      End
      Begin VB.Menu mnuproduct 
         Caption         =   "&Product"
      End
      Begin VB.Menu mnurapport 
         Caption         =   "&Rapport"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuverkoop 
         Caption         =   "&Verkoop"
      End
      Begin VB.Menu mnuvoorschrift 
         Caption         =   "V&oorschrift"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuinfo 
         Caption         =   "&Info"
      End
      Begin VB.Menu mnuinhoud 
         Caption         =   "I&nhoud"
      End
   End
End
Attribute VB_Name = "mdiapotheek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MDIForm_Load()
frmstart.Show
App.HelpFile = App.Path & "\" & "HELP.HLP"
End Sub

Private Sub mnuafdrukken_Click()

End Sub

Private Sub mnuarts_Click()
frmarts.Show
End Sub

Private Sub mnubestel_Click()
frmbestel.Show
End Sub

Private Sub mnuclose_Click()
Unload Me.ActiveForm
End Sub

Private Sub mnufactuur_Click()
frmfactuur.Show
End Sub

Private Sub mnuinfo_Click()
frminfo.Show
End Sub

Private Sub mnuinhoud_Click()

End Sub

Private Sub mnuklant_Click()
frmklant.Show
End Sub

Private Sub mnuleverancier_Click()
frmleverancier.Show
End Sub

Private Sub mnuopen_Click()

End Sub

Private Sub mnuopslaan_Click()

End Sub

Private Sub mnuproduct_Click()
frmproduct.Show
End Sub

Private Sub mnurapport_Click()
frmrapport.Show
End Sub

Private Sub mnustart_Click()
frmstart.Show
End Sub

Private Sub mnuverkoop_Click()
frmverkoop.Show
End Sub

Private Sub mnuvoorschrift_Click()
frmvoorschrift.Show
End Sub

Private Sub Picture1_Click()

End Sub
