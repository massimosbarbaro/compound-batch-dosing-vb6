VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSE 
   Caption         =   "SE"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSeId 
      Height          =   285
      Left            =   2880
      TabIndex        =   114
      Text            =   "Text1"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   11
      Left            =   5760
      TabIndex        =   113
      Text            =   "Text1"
      Top             =   6120
      Width           =   495
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   11
      Left            =   4080
      TabIndex        =   112
      Top             =   6360
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   111
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   9
      Left            =   4080
      TabIndex        =   110
      Top             =   5640
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   8
      Left            =   4080
      TabIndex        =   109
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   7
      Left            =   4080
      TabIndex        =   108
      Top             =   4920
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   107
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   106
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   105
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   104
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   103
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   102
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkS 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   101
      Top             =   2400
      Width           =   255
   End
   Begin VB.Frame fraIQ 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Indice di Qualità"
      Height          =   615
      Left            =   3840
      TabIndex        =   91
      Top             =   6960
      Visible         =   0   'False
      Width           =   2175
      Begin VB.OptionButton optIq 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   93
         Top             =   240
         Width           =   555
      End
      Begin VB.OptionButton optIq 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   92
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraEsterno 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Esterno"
      Height          =   615
      Left            =   5880
      TabIndex        =   88
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
      Begin VB.OptionButton optExt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   90
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optExt 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   89
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox txtDataRisposta 
      Height          =   285
      Left            =   5760
      TabIndex        =   87
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAzioni 
      Height          =   375
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   86
      Top             =   6960
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtRisposta 
      Height          =   375
      Left            =   6360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   85
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtNumeroRisposta 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5880
      TabIndex        =   84
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fraSoluzione 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Soluzione"
      Height          =   1095
      Left            =   4680
      TabIndex        =   80
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.OptionButton optSol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Respinto"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   83
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optSol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Accettato"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   82
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optSol 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Segnalazione"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   81
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdDettagli 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Dettagli"
      Height          =   615
      Left            =   6360
      MousePointer    =   1  'Arrow
      Picture         =   "frmSE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   645
      Index           =   3
      Left            =   5760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   78
      Text            =   "frmSE.frx":0532
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbFamiglia 
      DataSource      =   "datDifetto"
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   77
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbDifetto 
      DataSource      =   "datDifetto"
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   76
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtQtaAccettata 
      Height          =   285
      Left            =   5640
      TabIndex        =   74
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   73
      Text            =   "Text1"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   72
      Text            =   "Text1"
      Top             =   2400
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   2760
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   67
      Text            =   "Text1"
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   3120
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   3480
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   3840
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   4200
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   4560
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   6
      Left            =   2280
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   4920
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   7
      Left            =   2280
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   4920
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   8
      Left            =   480
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   8
      Left            =   1320
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   5280
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   8
      Left            =   2280
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   9
      Left            =   480
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   9
      Left            =   1320
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   5640
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   9
      Left            =   2280
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   10
      Left            =   480
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   10
      Left            =   1320
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   6000
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   10
      Left            =   2280
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   11
      Left            =   480
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   11
      Left            =   1320
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   6360
      Width           =   850
   End
   Begin VB.TextBox txtMPP 
      Height          =   285
      Index           =   11
      Left            =   2280
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   6360
      Width           =   615
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   2400
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   2760
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   3120
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   3480
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   3840
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   4200
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   6
      Left            =   3000
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   4560
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   7
      Left            =   3000
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   4920
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   5280
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   9
      Left            =   3000
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   5640
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   10
      Left            =   3000
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   6000
      Width           =   850
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   11
      Left            =   3000
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   6360
      Width           =   850
   End
   Begin VB.TextBox txtTot 
      Height          =   285
      Left            =   3120
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox txtId 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtId 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   0
      Left            =   5760
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   6
      Left            =   4920
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   8
      Left            =   4920
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   9
      Left            =   4920
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   10
      Left            =   4920
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtCMP 
      Height          =   285
      Index           =   11
      Left            =   4920
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   3
      Left            =   5760
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   4
      Left            =   5760
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   5
      Left            =   5760
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   6
      Left            =   5760
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   7
      Left            =   5760
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   8
      Left            =   5760
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   9
      Left            =   5760
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtCMPQ 
      Height          =   285
      Index           =   10
      Left            =   5760
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5760
      Width           =   495
   End
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   4800
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSE.frx":0538
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSE.frx":0694
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSE.frx":0BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSE.frx":111C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSE.frx":1660
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSE.frx":1BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSE.frx":20E8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   75
      Top             =   7815
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuovo"
            Key             =   "N"
            Object.ToolTipText     =   "Nuovo Reclamo"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "prova"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salva"
            Key             =   "S"
            Object.ToolTipText     =   "Salva Reclamo Corrente"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Annulla"
            Key             =   "A"
            Object.ToolTipText     =   "Annulla modifiche su reclamo corrente"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modifica"
            Key             =   "M"
            Object.ToolTipText     =   "Modifica il Reclamo Corrente"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Chiudi"
            Key             =   "C"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stampa"
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "phRight"
            Style           =   4
            Object.Width           =   500
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esci"
            Key             =   "X"
            Object.ToolTipText     =   "Esci"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   120
      Picture         =   "frmSE.frx":262C
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Id da produrre"
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
      TabIndex        =   100
      Top             =   120
      Width           =   3090
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label12 
      Caption         =   "ID"
      Height          =   255
      Left            =   840
      TabIndex        =   99
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "Data"
      Height          =   255
      Left            =   1560
      TabIndex        =   98
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "ODL"
      Height          =   255
      Left            =   2400
      TabIndex        =   97
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "Qta"
      Height          =   255
      Left            =   3240
      TabIndex        =   96
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "Totale"
      Height          =   255
      Left            =   2520
      TabIndex        =   95
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Mescola"
      Height          =   255
      Left            =   1680
      TabIndex        =   94
      Top             =   1080
      Width           =   615
   End
   Begin VB.Shape shpRisposta 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   4815
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   4335
   End
End
Attribute VB_Name = "frmSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String, rsATC As Recordset
Private fNuovoRecord As Boolean, sStato As String
Private mStatus As Integer
Private Enum CMDS
   cNuovo = 1
   cSalva = 3
   cAnnulla = 5
   cModifica = 7
   cChiudi = 9
   cStampa = 11
   cEsci = 13
End Enum
Private Enum Status
   cNew = 0
   cMod = 1
   cExists = 2
   cNotExists = 3
End Enum
Dim rsComponente As Recordset, rsN As Recordset
Dim fS As Boolean, dTot As Double
Public Property Let Status(vValue As Integer)
   If vValue >= cNew And vValue <= cNotExists Then
      mStatus = vValue
   End If
   With tbCommand
    If sFlag = "ATC" Then ' controllo che non possano modificare i dati dal MKG
      Select Case mStatus
         Case cNew, cMod
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = True
            .Buttons(cAnnulla).Enabled = True
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = False
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = True
            .Buttons(cChiudi).Enabled = (sStato = "S")
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
            fS = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = True
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
            fS = False
      End Select
    Else
      Select Case mStatus
         Case cNew, cMod
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
      End Select
    End If
    
    
    
   End With
'Private Sub NuovoReclamo()
'End Sub
End Property
Public Property Get Status() As Integer
   Status = mStatus
End Property

Public Property Let Key(vValue As String)
    mKey = Trim$(vValue)
'    CercaReclamoVenditore
'    CercaReclamoATC
    CercaID
   CercaComponenti
'    Set datElenco.Recordset = rsATC
'    datElenco.RecordSource = sqlc
'    txtODLInputTest.Text = mKey
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub Rete()
   Dim lDiff As Long
    Load frmAggiunte
    With frmAggiunte
        .Top = Me.Top + (Me.Height - frmAggiunte.Height)
        .Left = Me.Left + Me.Width
'   Controllo di non superare le dimensioni dello schermo
'       altrimenti la sposto indietro
        lDiff = Screen.Width * 0.98 - (.Left + .Width)
        If lDiff < 0 Then
            .Left = .Left + lDiff
        End If
        .Key = Me.Key
        .Show vbModal
    End With
End Sub

Private Sub chkS_Click(Index As Integer)
Dim i As Integer, sqlc As String
i = Index
If chkS(i).Value = 1 Then

    dTot = dTot + txtLotto(i).Text
'        sqlc = "Update Formule set SeID = 'P' " & _
'                "Where ID = '" & txtMP(i).Text & "' and " & _
'                "Caricata = 'S'"
'        db.Execute (sqlc)
Else
    dTot = dTot
End If
    txtTot.Text = dTot
End Sub
Private Sub SalvaMP()
Dim sqlc As String, rsMp As Recordset, i As Integer
'    sqlc = "Select Caricata, Id " & _
            "From Formule " & _
            " where ID='" & Left$(Me.Key, 5) & "'"
    
    sqlc = "SELECT SE.SeId, SE.SeS, SE.Data, SE.Qta " & _
            "FROM SE "
    
    Set rsMp = db.OpenRecordset(sqlc, dbOpenDynaset)
    
If fNuovoRecord Then
    rsMp.AddNew
        txtSeId.Text = CalcUltimaRisposta()
        txtSeId.Visible = True
        rsMp("SeId") = txtSeId.Text
        rsMp("Qta") = txtTot.Text
        rsMp("SeS") = "D"
        rsMp("Data") = Now()
'    rsmp.Update
Else
    rsMp.Edit
'       posso modificare solo la qta
       rsMp("Qta") = txtTot.Text
End If
    rsMp.Update
        
'Inserisco l'SeId nel db formule nel record id

'sqlc = "Select Id, SeId " &
'     "From Formule "
If fNuovoRecord Then
    rsN.MoveFirst
    For i = 0 To rsN.RecordCount - 1
        If chkS(i).Value = 1 Then
            rsN.Edit
            rsN("SeId") = txtSeId.Text
            rsN.Update
        End If
        rsN.MoveNext
    Next i
Else
    rsComponente.MoveFirst
    For i = 0 To rsComponente.RecordCount - 1
        If chkS(i).Value = 0 Then
            rsComponente.Edit
            rsComponente("SeId") = rsMp("SeId")
            rsComponente.Update
        End If
        rsComponente.MoveNext
    Next i

End If

    CercaComponenti
'     disabilito i campi
'   Campi False
   fNuovoRecord = False
End Sub



Private Sub SalvaRec()
   Dim scod As Variant
   If fNuovoRecord Then
      rsATC.AddNew
'      txtNumeroRisposta.Text = CalcUltimaRisposta()
      rsATC("NumeroRisposta") = txtNumeroRisposta.Text
      rsATC("ODL") = Me.Key
   Else
      rsATC.Edit
   End If
'   If Not IsNull(txtNumeroRisposta.Text) Then
'        rsATC("NumeroRisposta") = txtNumeroRisposta.Text
'   Else
'         rsATC("NumeroRisposta") = Null
'   End If
   If Not (IsNull(txtDataRisposta.Text) Or txtDataRisposta.Text = "") Then
      rsATC("DataRisposta") = txtDataRisposta.Text
   Else
       rsATC("DataRisposta") = Null
   End If
   scod = SalvaCodiceDifetto(cmbDifetto.Text)
'   lblCodDifetto.Caption = sCod
'   If Not (IsNull(lblCodDifetto.Caption) Or lblCodDifetto.Caption = "") Then
'       rsATC("CodDifetto") = lblCodDifetto.Caption
'   Else
'       rsATC("CodDifetto") = Null
'   End If
   If Not (IsNull(txtQtaAccettata.Text) Or txtQtaAccettata.Text = "") Then
        rsATC("QtaAccettata") = Val(txtQtaAccettata.Text)
   Else
         rsATC("QtaAccettata") = Null
   End If
   
'     Carico i dati del codfamiglia dall'rs
   scod = SalvaCodiceFamiglia(cmbFamiglia.Text)
'   lblCodFamiglia = sCod
'   If Not (IsNull(lblCodFamiglia.Caption) Or lblCodFamiglia.Caption = "") Then
'       rsATC("CodFamiglia") = lblCodFamiglia
'   Else
'       rsATC("CodFamiglia") = Null
'   End If
   If Not (IsNull(txtRisposta.Text) Or txtRisposta.Text = "") Then
       rsATC("Risposta") = txtRisposta
   Else
       rsATC("Risposta") = Null
   End If
   If Not (IsNull(txtAzioni.Text) Or txtAzioni.Text = "") Then
       rsATC("AzioniCorrettive") = txtAzioni.Text
   Else
       rsATC("AzioniCorrettive") = Null
   End If
   If optIq(0).Value Then
       rsATC("IndiceQualita") = True
   Else
       If optIq(1).Value Then
           rsATC("IndiceQualita") = False
       Else
           rsATC("IndiceQualita") = Null
       End If
   End If
   If optExt(0).Value Then
       rsATC("Esterno") = True
   Else
       If optExt(1).Value Then
           rsATC("Esterno") = False
       Else
           rsATC("Esterno") = Null
       End If
   End If
   If optSol(0).Value Then
       rsATC("Soluzione") = 0
   Else
       If optSol(1).Value Then
           rsATC("Soluzione") = 1
       Else
           If optSol(2).Value Then
              rsATC("Soluzione") = 2
           Else
   '            If optSol(3).Value Then
    '               rsATC("Soluzione") = 3
     '          Else
                   rsATC("Soluzione") = Null
       '        End If
          End If
       End If
   End If
   
   rsATC.Update
   CercaReclamoATC
'     disabilito i campi
   Campi False
   fNuovoRecord = False
End Sub

Private Sub Form_Load()
   Dim btn As Button
'  Carica i combo
'   ListCodiceDifetto
'   ListCodiceFamiglia
  fS = False
   Campi False
    dTot = 0
    'Nascondo tutti i pulsanti
'    cmdNew.Visible = False
'    cmdSalva.Visible = False
'    cmdAnnulla.Visible = False
'    cmdModifica.Visible = False
   For Each btn In tbCommand.Buttons
     If btn.Key <> "X" Then
        btn.Enabled = False
     End If
   Next
End Sub

Private Sub Form_Resize()
   Dim b As Button, l As Long, idx As Integer
   If Me.WindowState <> vbMinimized Then
      l = shpRisposta.Left * 2 + shpRisposta.Width + Me.Width - Me.ScaleWidth
      If Me.Width < l Then
         Me.Width = l
      End If
      l = 0
      Me.Refresh
      For Each b In tbCommand.Buttons
         If b.Caption <> "phRight" Then
            l = l + b.Width
         Else
            idx = b.Index
         End If
      Next
      Set b = tbCommand.Buttons(idx)
      b.Width = Me.ScaleWidth - l
      Set b = Nothing
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Load frmVenditori
'    frmVenditori.Show
'   rsATC.Close
'   Set rsATC = Nothin
   
   rsComponente.Close
   Set rsComponente = Nothing

End Sub

Private Sub CercaReclamoVenditore()
   Dim rsvenditore  As Recordset, sqlc As String, i As Integer
   
   sqlc = "SELECT ODL, DataReclamo, QtaReclamata, " & _
           "CausaReclamo, Venditore, NumeroReclamo, Stato " & _
           "FROM ReclamoVenditore " & _
           "where ODL='" & Me.Key & "'"
   Set rsvenditore = db.OpenRecordset(sqlc, dbOpenDynaset)
   If Not rsvenditore.EOF Then
     If sFlag = "ATC" Then ' Controllo chi apre il DB
      If rsvenditore("Stato") = "N" Then
         rsvenditore.Edit
         rsvenditore("Stato") = "V"
         rsvenditore.Update
         frmTV.ChangedTV = True
      End If
     End If
     
      For i = 0 To rsvenditore.Fields.Count - 1
         If Not IsNull(rsvenditore(i)) Then
             txtVenditore(i).Text = rsvenditore(i)
         Else
             txtVenditore(i).Text = ""
         End If
      Next i
      sStato = rsvenditore("Stato")
   Else
      Err.Raise 13, , "Errore nel data source"
   End If
   rsvenditore.Close
   Set rsvenditore = Nothing

End Sub
Private Sub CercaIDN()
Dim sqlc As String, i As Integer
'    sqlc = "SELECT Formule.ID, Formule.Mescola, MescPercCLP.MP, " & _
                "MescPercCLP.Per, DescrizMesc.Descrizione " & _
            "FROM (MescPercCLP RIGHT JOIN Formule ON " & _
                "MescPercCLP.Mescola = Formule.Mescola) " & _
                "LEFT JOIN DescrizMesc ON MescPercCLP.MP = DescrizMesc.MP " & _
            "GROUP BY Formule.ID, Formule.Mescola, MescPercCLP.MP, " & _
                "MescPercCLP.Per, DescrizMesc.Descrizione " & _
            "HAVING Formule.ID='" & Left$(Me.Key, 5) & "'"
'    If Not fNuovoRecord Then
'        sqlc = "SELECT F.Mescola, F.Data, F.OdL, F.ID, F.CodArt, " & _
                    "F.Qta, F.Caricata, F.SeID, SE.SeS, SE.Qta " & _
                "FROM Formule AS F INNER JOIN SE ON F.SeID = SE.SeId " & _
                "WHERE F.Mescola='" & Right$(Me.Key, 4) & "' AND SE.SeS='D'"
'    Else
        sqlc = "SELECT Mescola, Data, OdL, ID, CodArt, Qta, Caricata, SeID " & _
                "From Formule " & _
                "Where Mescola='" & LCase(Right$(Me.Key, 4)) & "' AND Caricata='s'"
'    End If


            i = 0
Set rsN = db.OpenRecordset(sqlc, dbOpenDynaset)
           Do While Not rsN.EOF
 '           For i = 0 To rsN.RecordCount - 1

                
                If Not IsNull(rsN("ID")) Then
                    txtMP(i).Visible = True
                    txtMP(i).Text = rsN("Id")
                    chkS(i).Visible = True
                Else
                    txtMP(i) = ""
                End If
                If Not IsNull(rsN("ODL")) Then
                    txtMPP(i).Visible = True
                    txtMPP(i).Text = rsN("ODL")
                Else
                    txtMPP(i).Text = ""
                End If
                If Not IsNull(rsN("Data")) Then
                    txtMPD(i).Visible = True
                    txtMPD(i).Text = rsN("Data")
                Else
                    txtMPD(i).Text = ""
                End If
                
                If Not IsNull(rsN("Qta")) Then
                    txtLotto(i).Visible = True
                    txtLotto(i).Text = rsN("Qta")
                Else
                    txtLotto(i).Text = ""
                End If
                i = i + 1
        rsN.MoveNext
        Loop
'        Next i
     
'   txtTot.Text = dTot







End Sub
Private Sub CercaComponenti()
Dim sqlc As String, i As Integer, dTot As Double

'    sqlc = "SELECT Formule.ID, Formule.Mescola, MescPercCLP.MP, " & _
                "MescPercCLP.Per, DescrizMesc.Descrizione " & _
            "FROM (MescPercCLP RIGHT JOIN Formule ON " & _
                "MescPercCLP.Mescola = Formule.Mescola) " & _
                "LEFT JOIN DescrizMesc ON MescPercCLP.MP = DescrizMesc.MP " & _
            "GROUP BY Formule.ID, Formule.Mescola, MescPercCLP.MP, " & _
                "MescPercCLP.Per, DescrizMesc.Descrizione " & _
            "HAVING Formule.ID='" & Left$(Me.Key, 5) & "'"
'    sqlc = "SELECT ID, MP, Lotto " & _
            "From MP " & _
            "WHERE ID='" & Left$(Me.Key, 5) & "'"
            
 'If Not fNuovoRecord Then
'    sqlc = "SELECT Mescola, Data, OdL, ID, CodArt, Qta, Caricata, SeID " & _
            "From Formule " & _
            "Where Mescola='" & Right$(Me.Key, 4) & "' AND Caricata='s'"
    
    
'    sqlc = "SELECT ID, OdL, Mescola, Qta, Caricata, SeID " & _
            "From Formule " & _
            "WHERE Mescola='" & Me.Key & "' AND Caricata='s'"
'Else
'    sqlc = "SELECT MP.MP, DescrizMesc.Descrizione, MP.Qta, MP.Lotto " & _
            "FROM (Formule INNER JOIN MP ON Formule.ID = MP.ID) " & _
                "LEFT JOIN DescrizMesc ON MP.MP = DescrizMesc.MP " & _
            "WHERE MP.ID='" & Left$(Me.Key, 5) & "'"

'End If
'    If Not fNuovoRecord Then
'        sqlc = "SELECT F.Mescola, F.Data, F.OdL, F.ID, F.CodArt, " & _
                    "F.Qta, F.Caricata, F.SeID, SE.SeS, SE.Qta " & _
                "FROM Formule AS F LEFT JOIN SE ON F.SeID = SE.SeId " & _
                "WHERE F.Mescola='" & Right$(Me.Key, 4) & "' AND F.Caricata='S' " & _
                    "AND ((SE.SeS)='D' Or (SE.SeS) Is Null)"
        sqlc = "SELECT F.Mescola, F.Data, F.OdL, F.ID, F.CodArt, " & _
                    "F.Qta, F.Caricata, F.SeID, SE.SeS, SE.Qta " & _
                "FROM Formule AS F LEFT JOIN SE ON F.SeID = SE.SeId " & _
                "WHERE F.Mescola='" & Right$(Me.Key, 4) & "' AND F.Caricata='S' " & _
                    "AND (SE.SeS)='D' "

'    Else
'        sqlc = "SELECT Mescola, Data, OdL, ID, CodArt, Qta, Caricata, SeID " & _
                "From Formule " & _
                "Where Mescola='" & Right$(Me.Key, 4) & "' AND Caricata='s'"
'    End If
    Set rsComponente = db.OpenRecordset(sqlc, dbOpenDynaset)
'    Do While Not rsComponente.EOF
    If Not rsComponente.EOF Then
        rsComponente.MoveLast
        
        i = rsComponente.RecordCount
        rsComponente.MoveFirst
        
        FoundID
        Me.Status = cExists
    Else
        Me.Status = cNotExists
    End If
'Loop
End Sub
Private Sub FoundID()
Dim i As Integer, dTot As Double

        For i = 0 To rsComponente.RecordCount - 1
            If Not IsNull(rsComponente("ID")) Then
                txtMP(i).Visible = True
                txtMP(i).Text = rsComponente("Id")
                chkS(i).Visible = False
            Else
                txtMP(i) = ""
            End If
            If Not IsNull(rsComponente("ODL")) Then
                txtMPP(i).Visible = True
                txtMPP(i).Text = rsComponente("ODL")
            Else
                txtMPP(i).Text = ""
            End If
            'Controllo se già memorizzato
'            If Not fNuovoRecord Then
            If Not IsNull(rsComponente("Data")) Then
                txtMPD(i).Visible = True
                txtMPD(i).Text = rsComponente("Data")
            Else
                txtMPD(i).Text = ""
            End If
            
            If Not IsNull(rsComponente("F.Qta")) Then
                txtLotto(i).Visible = True
                txtLotto(i).Text = rsComponente("F.Qta")
            Else
                txtLotto(i).Text = ""
            End If
            If Not IsNull(rsComponente("SeId")) Then
                chkS(i).Value = 1
            
            Else
                chkS(i).Value = 0
            End If
            If Not IsNull(rsComponente("Se.Qta")) Then
                txtTot.Text = rsComponente("Se.Qta")
            Else
                txtTot.Text = 0
            End If
            
        rsComponente.MoveNext
        Next i
     
 '  txtTot.Text = dTot
    
End Sub
Private Sub CercaID()
Dim sqlc As String, rsId As Recordset, i As Integer

sqlc = "SELECT Mescola, Caricata " & _
        "From Formule " & _
        "WHERE Mescola='" & Right$(Me.Key, 4) & "'"
        
Set rsId = db.OpenRecordset(sqlc, dbOpenDynaset)
        If Not rsId.EOF Then
        '    txtId(i).Text = rsId(i)
           For i = 0 To rsId.Fields.Count - 1
              If Not IsNull(rsId(i)) Then
                  txtId(i).Text = rsId(i)
              Else
                  txtId(i).Text = ""
              End If
           Next i
             sStato = rsId("Caricata")
        Else
           Err.Raise 13, , "Errore nel data source"
        End If

End Sub


Private Sub CercaReclamoATC()
    Dim sqlc As String, i As Integer
   
    sqlc = "SELECT ODL, NumeroRisposta, DataRisposta, " & _
            "Risposta, AzioniCorrettive, CodFamiglia, " & _
            "CodDifetto, Esterno, StatoReclamo, IndiceQualita, " & _
            "Soluzione, QtaAccettata FROM ATC where ODL='" & Me.Key & "'"
    Set rsATC = db.OpenRecordset(sqlc, dbOpenDynaset)
    
   If Not rsATC.EOF Then
      FoundODL
      Me.Status = cExists
    Else
      Me.Status = cNotExists
    End If
    'For i = 0 To rsATC.Fields.Count - 1
    '    txtReclami(i).DataField = rsATC.Fields(i).Name
    'Next i
End Sub
Private Sub FoundODL()
   Dim iSoluz As Integer
    If Not IsNull(rsATC("NumeroRisposta")) Then
        txtNumeroRisposta = rsATC("NumeroRisposta")
    Else
        txtNumeroRisposta = ""
    End If
    If Not IsNull(rsATC("DataRisposta")) Then
        txtDataRisposta = rsATC("DataRisposta")
    Else
        txtDataRisposta = ""
    End If
    If Not IsNull(rsATC("CodDifetto")) Then
'        lblCodDifetto = rsATC("CodDifetto")
'        cmbDifetto.Text = FoundCodiceDifetto(lblCodDifetto.Caption)
    Else
'        lblCodDifetto = ""
        cmbDifetto.ListIndex = -1
    End If
    If Not IsNull(rsATC("CodFamiglia")) Then
      'lblCodFamiglia = rsATC("CodFamiglia")
'      cmbFamiglia.Text = FoundCodiceFamiglia(lblCodFamiglia.Caption)
    Else
'     lblCodFamiglia = ""
     cmbFamiglia.ListIndex = -1
    End If
    If Not IsNull(rsATC("Risposta")) Then
      txtRisposta = rsATC("Risposta")
    Else
        txtRisposta = ""
    End If
    If Not IsNull(rsATC("AzioniCorrettive")) Then
        txtAzioni = rsATC("AzioniCorrettive")
    Else
        txtAzioni = ""
    End If
    If Not IsNull(rsATC("IndiceQualita")) Then
        If rsATC("IndiceQualita") Then
            optIq(0).Value = True
        Else
            optIq(1).Value = True
        End If
    Else
          optIq(0).Value = False
          optIq(1).Value = False
    End If
    If Not IsNull(rsATC("Esterno")) Then
        If rsATC("Esterno") Then
            optExt(0).Value = True
        Else
            optExt(1).Value = True
        End If
    Else
        optExt(0).Value = False
        optExt(1).Value = False
    End If
    If Not IsNull(rsATC("Soluzione")) Then
      iSoluz = rsATC("Soluzione")
      optSol(iSoluz).Value = True
    Else
        optSol(0).Value = False
        optSol(1).Value = False
        optSol(2).Value = False
    End If
   If Not IsNull(rsATC("QtaAccettata")) Then
      txtQtaAccettata = rsATC("QtaAccettata")
    Else
        txtQtaAccettata = ""
    End If
End Sub
Private Sub ListCodiceDifetto()
   Dim rsDifetto As Recordset, sqlc As String, sMsg As String
   sqlc = "SELECT DescrizioneDifetto " & _
           "FROM Difetto "
   Set rsDifetto = db.OpenRecordset(sqlc, dbOpenSnapshot)
   cmbDifetto.Clear
   Do While Not rsDifetto.EOF
      cmbDifetto.AddItem rsDifetto(0)
      rsDifetto.MoveNext
   Loop
   rsDifetto.Close
   Set rsDifetto = Nothing
   cmbDifetto.ListIndex = -1
End Sub
Private Sub ListCodiceFamiglia()
    Dim rsFamiglia As Recordset, sqlc As String, sMsg As String
    sqlc = "SELECT DescrizioneFamiglia " & _
            "FROM Famiglia order by 1"
    Set rsFamiglia = db.OpenRecordset(sqlc, dbOpenSnapshot)
    cmbFamiglia.Clear
    Do While Not rsFamiglia.EOF
        cmbFamiglia.AddItem rsFamiglia(0)
        rsFamiglia.MoveNext
    Loop
    rsFamiglia.Close
    Set rsFamiglia = Nothing
    cmbFamiglia.ListIndex = -1
    
End Sub
Private Function FoundCodiceDifetto(cod As String)
   Dim rsDifetto As Recordset, sqlc As String
   sqlc = "SELECT DescrizioneDifetto " & _
           "FROM Difetto where CodDifetto='" & cod & "'"
   Set rsDifetto = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsDifetto.EOF Then
       FoundCodiceDifetto = rsDifetto("DescrizioneDifetto")
   Else
       FoundCodiceDifetto = ""
   End If
   rsDifetto.Close
   Set rsDifetto = Nothing
End Function
Private Function FoundCodiceFamiglia(cod As String)
   Dim rsFamiglia As Recordset, sqlc As String
   sqlc = "SELECT DescrizioneFamiglia " & _
           "FROM Famiglia where CodFamiglia='" & cod & "'"
   Set rsFamiglia = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsFamiglia.EOF Then
       FoundCodiceFamiglia = rsFamiglia("DescrizioneFamiglia")
   Else
       FoundCodiceFamiglia = ""
   End If
   rsFamiglia.Close
   Set rsFamiglia = Nothing
End Function
Private Function SalvaCodiceDifetto(s As String)
   Dim rsDifetto As Recordset, sqlc As String, sDescrizione As String
   Dim iPos As Integer
   iPos = InStr(s, "'")
   If iPos <> 0 Then
      s = Left$(s, iPos) & "'" & Right$(s, Len(s) - iPos)
      
   End If
   sqlc = "SELECT CodDifetto " & _
           "FROM Difetto where DescrizioneDifetto='" & Trim$(s) & "'"
   Set rsDifetto = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsDifetto.EOF Then
      SalvaCodiceDifetto = rsDifetto("CodDifetto")
   Else
      SalvaCodiceDifetto = ""
   End If
   rsDifetto.Close
   Set rsDifetto = Nothing
End Function

Private Function SalvaCodiceFamiglia(s As String)
   Dim rsFamiglia As Recordset, sqlc As String, sDescrizione As String
   
   sqlc = "SELECT CodFamiglia " & _
           "FROM Famiglia where DescrizioneFamiglia='" & s & "'"
   Set rsFamiglia = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsFamiglia.EOF Then
       SalvaCodiceFamiglia = rsFamiglia("CodFamiglia")
   Else
      SalvaCodiceFamiglia = ""
   End If
   rsFamiglia.Close
   Set rsFamiglia = Nothing
End Function

Private Sub Campi(Flag As Boolean)
    Dim i As Integer, dr As Double
    
    txtNumeroRisposta.Enabled = False
    txtDataRisposta.Enabled = Flag
    txtRisposta.Enabled = Flag
    txtAzioni.Enabled = Flag
    fraIQ.Enabled = Flag
    fraEsterno.Enabled = Flag
    fraSoluzione.Enabled = Flag
    cmbDifetto.Enabled = Flag
    cmbFamiglia.Enabled = Flag
    txtQtaAccettata.Enabled = Flag
    txtNumeroRisposta.Locked = Not Flag
    txtDataRisposta.Locked = Not Flag
    txtRisposta.Locked = Not Flag
    txtAzioni.Locked = Not Flag
    txtQtaAccettata.Locked = Not Flag
    If fS Then
        dr = rsComponente.RecordCount - 1
    Else
        dr = txtMP.Count - 1
    End If
    For i = 0 To dr
    txtMP(i).Visible = Flag
        txtMPD(i).Visible = Flag
        txtMPP(i).Visible = Flag
        txtLotto(i).Visible = Flag
        txtCMP(i).Visible = Flag
        txtCMPQ(i).Visible = Flag
        chkS(i).Visible = Flag
    Next i
    
    txtSeId.Visible = False
    
End Sub

Private Sub CancellaCampi()
   txtNumeroRisposta.Text = ""
   txtDataRisposta.Text = ""
   txtRisposta.Text = ""
   txtAzioni.Text = ""
   txtQtaAccettata.Text = ""
   optIq(0).Value = False
   optIq(1).Value = False
   
   optExt(0).Value = False
   optExt(1).Value = False
   
   optSol(0).Value = False
   optSol(1).Value = False
   optSol(2).Value = False
   
   cmbDifetto.ListIndex = -1
   cmbFamiglia.ListIndex = -1
End Sub

Private Sub tbCommand_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim iRes As Integer, rsTmp As Recordset, sqlc As String
   Select Case Button.Key
      Case "X"
         Unload Me
      Case "N"
         fNuovoRecord = True
'         Campi True
'         CancellaCampi
'            AttivaC
            CercaIDN
         Me.Status = cNew
         CampiD True
      Case "S"
'         SalvaRec
        SalvaMP
        Me.Status = cExists
'        Campi False
        CampiD False
      Case "A"
         If fNuovoRecord Then
            CancellaCampi
            Me.Status = cNotExists
         Else
'            FoundODL
'           FoundID
           CercaComponenti
            Me.Status = cExists
         End If
'         Campi False
        CampiD False
      Case "M"
         fNuovoRecord = False
         Campi True
         Me.Status = cMod
'         Controtipo
         CampiD True
      Case "C"
         iRes = MsgBox("Vuoi chiudere la produzione corrente ", _
               vbYesNo + vbDefaultButton2 + vbQuestion)
         If iRes = vbYes Then
            ChiudiSe
            
 '           sqlc = "Update Formule set Caricata = 'P' " & _
                     "Where ID = '" & Left$(Me.Key, 5) & "' and " & _
                     "Caricata = 'S'"
 '           db.Execute (sqlc)
 '           sqlc = "Update ATC set StatoReclamo = 'N' " & _
                  "where ODL = '" & txtVenditore(0) & "'"
 '           db.Execute (sqlc)
 '           sStato = "C"
            Me.Status = cExists
            frmTV.ChangedTV = True
         End If
      Case "P"
         fNuovoRecord = False
         Rete
         Me.Status = cExists
   End Select
End Sub

Private Sub Controtipo()
Dim sqlc As String, rsCon As Recordset
Dim sWhere As String, i As Integer

For i = 0 To rsComponente.RecordCount - 1
    sWhere = txtMP(i).Text
    sqlc = "SELECT MP, Controtipo " & _
            "From Controtipo " & _
            "WHERE MP='" & sWhere & "'"
    Set rsCon = db.OpenRecordset(sqlc, dbOpenDynaset)
        If Not rsCon.EOF Then
            txtCMP(i).Text = rsCon("Controtipo")
            txtCMPQ(i).Visible = True
 '           txtCMPQ(i).Text = ""
        Else
            txtCMP(i).Text = ""
            txtCMP(i).Visible = False
            txtCMPQ(i).Visible = False
        End If
            txtCMPQ(i).Text = ""

Next i

        
        
        
        
End Sub
Private Sub CampiD(Flag As Boolean)
Dim dr As Double, i As Integer
    If fS Then
        dr = rsComponente.RecordCount - 1
    Else
        dr = txtMP.Count - 1
    End If
    For i = 0 To dr
    txtMP(i).Enabled = Flag
        txtMPD(i).Enabled = Flag
        txtMPP(i).Enabled = Flag
        txtLotto(i).Enabled = Flag
        txtCMP(i).Enabled = Flag
        txtCMPQ(i).Enabled = Flag
        
    Next i
    
End Sub


Private Sub AttivaC()
Dim dr As Double, i As Integer
    If fS Then
        dr = rsComponente.RecordCount - 1
    Else
        dr = txtMP.Count - 1
    End If
    For i = 0 To dr
    chkS(i).Visible = True
            
    Next i
End Sub
Private Function CalcUltimaRisposta()
   Dim rsT As Recordset, sqlc As String
   sqlc = "update UltimaRisposta " & _
               "set UltimaRisposta = UltimaRisposta + 1 " & _
               "Where dummy=0"
   db.Execute (sqlc)
   sqlc = "Select UltimaRisposta from UltimaRisposta " & _
               "Where dummy = 0"
   Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsT.EOF Then
      CalcUltimaRisposta = rsT(0)
   Else
      Err.Raise 555, , "Errore nel calcolo del numero risposta"
   End If
   rsT.Close
   Set rsT = Nothing
End Function
Private Sub ChiudiSe()
Dim i As Integer
    rsComponente.MoveFirst
    For i = 0 To rsComponente.RecordCount - 1
        If chkS(i).Value = 1 Then
            rsComponente.Edit
            If IsNull(rsComponente("SeId")) Then
                rsComponente("SeId") = txtSeId.Text
            End If
            
            rsComponente("SeS") = "C"
            rsComponente("Caricata") = "P"
            rsComponente.Update
        End If
        rsComponente.MoveNext
    Next i

End Sub
