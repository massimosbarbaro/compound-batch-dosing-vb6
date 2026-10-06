VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNewOrd 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Scarico"
   ClientHeight    =   4680
   ClientLeft      =   5130
   ClientTop       =   3735
   ClientWidth     =   5055
   Icon            =   "frmNewOrd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   5055
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   615
      Left            =   4320
      Picture         =   "frmNewOrd.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "Iporta"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ComboBox cmbMesi 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSMask.MaskEdBox mskDataDA 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataA 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuovi ordini"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   630
      Left            =   1200
      TabIndex        =   8
      Top             =   240
      Width           =   2775
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmNewOrd.frx":0E14
      Top             =   0
      Width           =   450
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Giorno/Mese / Anno (4 cifre)"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "A"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Da"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   4455
   End
End
Attribute VB_Name = "frmNewOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iIndice As Integer
'Dim ac As Access.Application

Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub cmdGrafico_Click()
   Dim sSelFormula As String, sGrafico As String
   Dim ac As Access.Application, sReport As String
   Dim sFile As String, sWhere As String, sTabella As String

'-------------------------------------------------------------------------------
' Preparo le date limite per la stampa dei grafici

   Dim sqlc As String, rsLimiti As Recordset, i As Byte
   Dim dDataInizio, dDataFine
   Dim sTitolo As String, sTest As String, sLinea As String, sGraficoFamiglia As String, sFam As String
   Dim sReteGrafico As String, sQtaGrafico As String, sGraficiReti As String
   
   
'   sDataDa = "00.00.00"
'   sDataA = "00.00.00"
   
   If mskDataA.Text = "__/__/____" Then
      mskDataA.Text = mskDataDA.Text
   End If
   dDataFine = Format((mskDataA.Text), "mm/dd/yyyy")
   
'   sTest = Format((mskDataDA.Text), "dd/mm/yyyy")
   dDataInizio = Format((mskDataDA.Text), "mm/dd/yyyy")
'   dDataInizio = Format(sTest, "yyyy/mm/dd")
'    DataUsa (dDataInizio)
 '   DataUsa (dDataFine)
   If dDataFine < dDataInizio Then
      MsgBox "Date Errate !!"
      Exit Sub
   End If
   
   sDataDa = dDataInizio
   sDataA = dDataFine
   
   
'Fine delle date
        Load frmExportMFG
        frmExportMFG.Key = "Aggiorna"
        frmExportMFG.Show vbModal

'esco
Unload Me


End Sub

Private Sub Form_Load()
iIndice = -99
   Dim i As Byte
   For i = 1 To 12
      cmbMesi.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
   Next i
   cmbMesi.Text = Format(Month(Now), "mmmm")
   mskDataDA.Text = Format(Now, "dd/mm/yyyy")
   Me.Show
'   mskDataDA.SetFocus
'    optGrafico(0) = True
End Sub

Private Sub mskDataA_Validate(Cancel As Boolean)
   If mskDataA.Text <> "__/__/____" Then
      If Not IsDate(mskDataA.Text) Then
         MsgBox "Data A Non valida"
         Cancel = True
      End If
   Else
      mskDataA.Text = mskDataDA.Text
   End If
End Sub

Private Sub mskDataDA_Validate(Cancel As Boolean)
'  La DATADA NON può essere nulla !!
'   If mskDataDA.Text <> "__/____" Then
   If Not IsDate(mskDataDA.Text) Then
      MsgBox "Data DA Non valida"
      Cancel = True
   End If
'   End If
End Sub


Private Sub optGrafico_Click(Index As Integer)
iIndice = Index

End Sub



