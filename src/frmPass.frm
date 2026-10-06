VERSION 5.00
Begin VB.Form frmPass 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Password"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   Icon            =   "frmPass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "Esci"
      Default         =   -1  'True
      Height          =   615
      Left            =   5640
      Picture         =   "frmPass.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtSE 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Verifica"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmPass.frx":0E14
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Password per le aggiunte"
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
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   5445
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label Label19 
      Caption         =   "Password"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Property Let Key(vValue As String)
    mKey = Trim$(vValue)
    txtSE.Text = mKey
'    CercaID
'    CercaComponenti
End Property
Public Property Get Key() As String
    Key = mKey
End Property

Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtPass.Text = "changeme" Then
        Load frmAggiunte
        frmAggiunte.Key = txtSE.Text
        frmAggiunte.Show
    Else
        MsgBox "La password inserita non è valida! Riprovate od uscite.", vbCritical
     End If
     
    
End Sub

Private Sub Form_Load()
txtSE.Visible = False
End Sub

Private Sub txtId_Change()

End Sub
