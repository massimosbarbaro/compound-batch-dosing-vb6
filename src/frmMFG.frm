VERSION 5.00
Begin VB.Form frmMFG 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dettagli - MFG"
   ClientHeight    =   6330
   ClientLeft      =   7065
   ClientTop       =   465
   ClientWidth     =   4800
   Icon            =   "frmMFG.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Cancel          =   -1  'True
      Caption         =   "Stampa"
      Default         =   -1  'True
      Height          =   615
      Left            =   3000
      Picture         =   "frmMFG.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5640
      Width           =   655
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   615
      Left            =   3960
      Picture         =   "frmMFG.frx":1334
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   12
      Left            =   1080
      TabIndex        =   31
      Text            =   "10500"
      Top             =   4035
      Width           =   615
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   14
      Left            =   3600
      TabIndex        =   30
      Text            =   "Besafilm"
      Top             =   4035
      Width           =   855
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   29
      Text            =   "119"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   7
      Left            =   4080
      TabIndex        =   28
      Text            =   "111"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   25
      Text            =   "1200"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   11
      Left            =   4080
      TabIndex        =   24
      Text            =   "3500"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   22
      Text            =   "1166"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   10
      Left            =   2520
      TabIndex        =   20
      Text            =   "1500"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   18
      Text            =   "22693.13"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   19
      Left            =   1920
      TabIndex        =   10
      Text            =   "10/12/2000"
      Top             =   4995
      Width           =   975
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   18
      Left            =   1200
      TabIndex        =   9
      Text            =   "C500"
      Top             =   4995
      Width           =   615
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   17
      Left            =   1920
      TabIndex        =   8
      Text            =   "10/12/2000"
      Top             =   4515
      Width           =   975
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   16
      Left            =   1200
      TabIndex        =   7
      Text            =   "BJ500"
      Top             =   4515
      Width           =   615
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   15
      Left            =   2400
      TabIndex        =   6
      Text            =   "p2"
      Top             =   4035
      Width           =   375
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   13
      Left            =   3720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   4
      Text            =   ""
      Top             =   2040
      Width           =   2915
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   4
      Left            =   300
      TabIndex        =   3
      Text            =   "0123456789012"
      Top             =   2040
      Width           =   1275
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   2
      Text            =   "ENERGOPOL-TRADE-POLNOC SP.ZO"
      Top             =   1215
      Width           =   3015
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   1
      Text            =   "22020110"
      Top             =   1215
      Width           =   855
   End
   Begin VB.TextBox txtDettagli 
      DataSource      =   "datMFG"
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   0
      Text            =   "SUD"
      Top             =   600
      Width           =   615
   End
   Begin VB.Image imgLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmMFG.frx":147E
      Top             =   0
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Produzione"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   3600
      Width           =   795
   End
   Begin VB.Shape Shape3 
      Height          =   1575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   4575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Esterno"
      Height          =   195
      Left            =   3000
      TabIndex        =   35
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Qta spedita"
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   4080
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Linea"
      Height          =   195
      Left            =   1920
      TabIndex        =   33
      Top             =   2565
      Width           =   390
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Colore"
      Height          =   195
      Left            =   3360
      TabIndex        =   32
      Top             =   2565
      Width           =   450
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   495
      Width           =   4215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Larghezza"
      Height          =   195
      Left            =   3240
      TabIndex        =   27
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Spessore"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   3000
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Finitura"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   2565
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Altezza"
      Height          =   195
      Left            =   1800
      TabIndex        =   21
      Top             =   3000
      Width           =   510
   End
   Begin VB.Label lblODL 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "ODL"
      Height          =   195
      Left            =   600
      TabIndex        =   19
      Top             =   645
      Width           =   330
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "ODV"
      Height          =   195
      Left            =   3240
      TabIndex        =   17
      Top             =   630
      Width           =   345
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ubic."
      Height          =   195
      Left            =   1920
      TabIndex        =   16
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prodotto"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Calandra"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   5040
      Width           =   630
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bobinatrice"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   4560
      Width           =   795
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Rete"
      Height          =   195
      Left            =   2040
      TabIndex        =   12
      Top             =   645
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cliente"
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Height          =   1455
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   4575
   End
End
Attribute VB_Name = "frmMFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private mKey As String


Public Property Let Key(ByVal vValue As String)
    mKey = Trim$(vValue)
    Dim sqlc As String, rsMFG As Recordset, i As Integer
    sqlc = "SELECT ODL, Rete, CodCli, Cliente," & _
            " CodProd, Prodotto, LineaProdotto, Colore," & _
            " Finitura, Spessore, Altezza, Larghezza," & _
            " Qta, ODV, Esterni, Ubicazione," & _
            " Bobinatrice, DataBobinatrice, Calandra, DataCalandra" & _
            " FROM MFG where ODL = '" & Me.Key & "'"
    Set rsMFG = db.OpenRecordset(sqlc, dbOpenSnapshot)
    For i = 0 To rsMFG.Fields.Count - 1
         If Not IsNull(rsMFG(i)) Then
                txtDettagli(i).Text = rsMFG(i)
            Else
                txtDettagli(i) = ""
            End If
    Next i
 
 
End Property
Public Property Get Key() As String
    Key = mKey
End Property

Private Sub cmdEsci_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Me.PrintForm

End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To Me.txtDettagli.Count - 1
        txtDettagli(i).Enabled = True
        txtDettagli(i).Locked = True
        
    Next i
End Sub

