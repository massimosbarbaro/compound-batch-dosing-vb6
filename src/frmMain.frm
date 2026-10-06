VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main"
   ClientHeight    =   5175
   ClientLeft      =   5115
   ClientTop       =   3210
   ClientWidth     =   4950
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4950
   Begin VB.CommandButton cmdGraf 
      Caption         =   "&Stampe"
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   615
      Left            =   4080
      Picture         =   "frmMain.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdStatistiche 
      Caption         =   "&Aggiorna"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   2460
      Width           =   1455
   End
   Begin VB.CommandButton cmdSegnalazioni 
      Caption         =   "&Import"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdReclami 
      Caption         =   "&Bilance"
      Height          =   615
      Left            =   1800
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   120
      Picture         =   "frmMain.frx":0E14
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Produzione"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2895
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Menu mnuReclami 
      Caption         =   "Reclami"
      Visible         =   0   'False
      Begin VB.Menu mnuR 
         Caption         =   "Produzione"
         Index           =   0
      End
      Begin VB.Menu mnuR 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "Nuovo Reclamo"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "Posta"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "Visualizza"
         Index           =   5
      End
      Begin VB.Menu mnuR 
         Caption         =   "aTC/mKG"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuR 
         Caption         =   "Marketing"
         Index           =   7
      End
   End
   Begin VB.Menu mnuStatistiche 
      Caption         =   "Statistiche"
      Visible         =   0   'False
      Begin VB.Menu mnuS 
         Caption         =   "Importa ordini"
         Index           =   0
      End
      Begin VB.Menu mnuS 
         Caption         =   "Importa formule"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsci_Click()
   Unload Me
End Sub

Private Sub cmdGraf_Click()
        Load frmGrafici
        frmGrafici.Show

End Sub

Private Sub cmdReclami_Click()
'    Load frmVenditori
'    frmVenditori.Show
End Sub

Private Sub cmdReclami_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case sFlag
      Case "ATC"
         mnuR(0).Visible = True
         mnuR(5).Visible = False
         mnuR(7).Visible = False
      Case "MKG"
         mnuR(0).Visible = False
         mnuR(5).Visible = True
         mnuR(7).Visible = False
      Case "ATC/MKG"
         mnuR(0).Visible = True
         mnuR(5).Visible = True
         mnuR(7).Visible = False
      Case "VIEW"
         mnuR(0).Visible = False
         mnuR(5).Visible = False
         mnuR(7).Visible = True
      Case "BIL"
         mnuR(0).Visible = False
         mnuR(5).Visible = True
         mnuR(7).Visible = False

      Case Else
         MsgBox "Questo messaggio non esiste !!"
   End Select
   PopupMenu mnuReclami, vbPopupMenuCenterAlign
End Sub


Private Sub cmdSegnalazioni_Click()
'   Load frmImportMFG
'   frmImportMFG.Show vbModal
End Sub

Private Sub cmdStatistiche_Click()
   Screen.MousePointer = vbArrowHourglass
   Screen.MousePointer = vbNormal
End Sub

Private Sub cmdStatistiche_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PopupMenu mnuStatistiche, vbPopupMenuCenterAlign
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer, res As Integer, nRs As Integer
    res = MsgBox("Esco dall'Applicazione ?", vbQuestion + vbYesNo + vbDefaultButton1, _
        "Reclami - Uscita")
    If res = vbYes Then
'
'   Check su altre forms caricate
'
        For i = Forms.Count - 1 To 0 Step -1
            If Forms(i).Name <> Me.Name Then
    '            Set f = Forms(i)
                Unload Forms(i)
            End If
        Next i
    '
    '   Controllo rs aperti !!!  (Ben Fatto)
    '
        nRs = db.Recordsets.Count
        For i = 1 To nRs
            db.Recordsets(0).Close
        Next i
    '
    '   Chiusura DB
    '
        If fDBOk Then
            db.Close
        End If
        Set db = Nothing
    Else
        Cancel = True
    End If
End Sub

Private Sub mnuR_Click(Index As Integer)
   Select Case Index
      Case 0
         Load frmTV
         frmTV.Param = "ATC"
         frmTV.Show
      Case 2
         Load frmReclamiVenditori
      Case 4
'         Load frmPop3
'         frmPop3.Show vbModal
      Case 5
         Load frmTV
         frmTV.Param = "MKG"
         frmTV.Show
'      Case 6
'         Load frmTV
'         frmTV.Param = "ATC/MKG"
'         frmTV.Show
      Case 7
         Load frmTV
         frmTV.Param = "VIEW"
         frmTV.Show
   
   End Select
End Sub

Private Sub mnuS_Click(Index As Integer)
Select Case Index
    Case 0
'        Load frmExportMFG
'        frmExportMFG.Key = "Aggiorna"
'        frmExportMFG.Show vbModal
        Load frmNewOrd
        frmNewOrd.Show
        
    Case 1
        Load frmExportMFG
        frmExportMFG.Key = "Formule"
        frmExportMFG.Show vbModal
End Select
End Sub
