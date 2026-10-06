VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmBilance 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Bilance"
   ClientHeight    =   9000
   ClientLeft      =   2880
   ClientTop       =   1185
   ClientWidth     =   9435
   Icon            =   "frmBilance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   9435
   Begin VB.TextBox txtMescole 
      Height          =   285
      Left            =   720
      TabIndex        =   131
      Text            =   "Text1"
      Top             =   7920
      Width           =   2055
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   13
      Left            =   6600
      TabIndex        =   130
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   4560
      TabIndex        =   129
      Top             =   7560
      Width           =   855
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   13
      Left            =   7920
      TabIndex        =   128
      Text            =   "Text1"
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   3480
      TabIndex        =   127
      Text            =   "Text1"
      Top             =   7560
      Width           =   855
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   13
      Left            =   1080
      TabIndex        =   126
      Text            =   "Text1"
      Top             =   7560
      Width           =   2145
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   125
      Text            =   "Text1"
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   13
      Left            =   5640
      TabIndex        =   124
      Text            =   "Text1"
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   12
      Left            =   6600
      TabIndex        =   123
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   4560
      TabIndex        =   122
      Top             =   7200
      Width           =   855
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   12
      Left            =   7920
      TabIndex        =   121
      Text            =   "Text1"
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   3480
      TabIndex        =   120
      Text            =   "Text1"
      Top             =   7200
      Width           =   855
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   12
      Left            =   1080
      TabIndex        =   119
      Text            =   "Text1"
      Top             =   7200
      Width           =   2145
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   12
      Left            =   240
      TabIndex        =   118
      Text            =   "Text1"
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   12
      Left            =   5640
      TabIndex        =   117
      Text            =   "Text1"
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txtDM 
      Height          =   285
      Left            =   1920
      TabIndex        =   116
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtParte 
      Height          =   285
      Left            =   5760
      TabIndex        =   114
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   112
      Text            =   "Text1"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   111
      Text            =   "Text1"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   110
      Text            =   "Text1"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   109
      Text            =   "Text1"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   4
      Left            =   5640
      TabIndex        =   108
      Text            =   "Text1"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   107
      Text            =   "Text1"
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   106
      Text            =   "Text1"
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   7
      Left            =   5640
      TabIndex        =   105
      Text            =   "Text1"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   8
      Left            =   5640
      TabIndex        =   104
      Text            =   "Text1"
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   9
      Left            =   5640
      TabIndex        =   103
      Text            =   "Text1"
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   10
      Left            =   5640
      TabIndex        =   102
      Text            =   "Text1"
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Index           =   11
      Left            =   5640
      TabIndex        =   101
      Text            =   "Text1"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox txtSPC 
      Height          =   285
      Left            =   5520
      TabIndex        =   100
      Text            =   "Text1"
      Top             =   7920
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   83
      Text            =   "Text1"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   82
      Text            =   "Text1"
      Top             =   2880
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   81
      Text            =   "Text1"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   80
      Text            =   "Text1"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   79
      Text            =   "Text1"
      Top             =   3240
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   78
      Text            =   "Text1"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   76
      Text            =   "Text1"
      Top             =   3600
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   75
      Text            =   "Text1"
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   74
      Text            =   "Text1"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   73
      Text            =   "Text1"
      Top             =   3960
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   72
      Text            =   "Text1"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   71
      Text            =   "Text1"
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   4320
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3480
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   67
      Text            =   "Text1"
      Top             =   4680
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3480
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   64
      Text            =   "Text1"
      Top             =   5040
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3480
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   62
      Text            =   "Text1"
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   5400
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   3480
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   5760
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   3480
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   6120
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   6120
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   3480
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   10
      Left            =   1080
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   6480
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   3480
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Index           =   11
      Left            =   1080
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   6840
      Width           =   2145
   End
   Begin VB.TextBox txtMPP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3480
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   0
      Left            =   7920
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   1
      Left            =   7920
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   2
      Left            =   7920
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   3
      Left            =   7920
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   4
      Left            =   7920
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   5
      Left            =   7920
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   6
      Left            =   7920
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   7
      Left            =   7920
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   8
      Left            =   7920
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   9
      Left            =   7920
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   10
      Left            =   7920
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtLotto 
      Height          =   285
      Index           =   11
      Left            =   7920
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtTot 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   7920
      Width           =   1095
   End
   Begin VB.TextBox txtMescola 
      Height          =   285
      Left            =   960
      TabIndex        =   34
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   7200
      TabIndex        =   33
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtODL 
      Height          =   525
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   960
      Width           =   8535
   End
   Begin VB.TextBox txtId 
      Height          =   525
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   1560
      Width           =   8535
   End
   Begin VB.TextBox txtQta 
      Height          =   285
      Left            =   1080
      TabIndex        =   30
      Text            =   "0"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtSeId 
      Height          =   285
      Left            =   8760
      TabIndex        =   29
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Left            =   3240
      TabIndex        =   28
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   27
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   26
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4560
      TabIndex        =   25
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   4560
      TabIndex        =   24
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4560
      TabIndex        =   23
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4560
      TabIndex        =   22
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   4560
      TabIndex        =   21
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   4560
      TabIndex        =   20
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   4560
      TabIndex        =   19
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   4560
      TabIndex        =   18
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   4560
      TabIndex        =   17
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txtFatt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   4560
      TabIndex        =   16
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox txtSF 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   7920
      Width           =   855
   End
   Begin VB.TextBox txtCicli 
      Height          =   285
      Left            =   4680
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdCalcola 
      Caption         =   "Calcola"
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   12
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   11
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   10
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   3
      Left            =   6600
      TabIndex        =   9
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   8
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   5
      Left            =   6600
      TabIndex        =   7
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   6
      Left            =   6600
      TabIndex        =   6
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   7
      Left            =   6600
      TabIndex        =   5
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   8
      Left            =   6600
      TabIndex        =   4
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   9
      Left            =   6600
      TabIndex        =   3
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   10
      Left            =   6600
      TabIndex        =   2
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox txtFT 
      Height          =   285
      Index           =   11
      Left            =   6600
      TabIndex        =   1
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox txtSFT 
      Height          =   285
      Left            =   6480
      TabIndex        =   0
      Top             =   7920
      Width           =   855
   End
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   8280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBilance.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBilance.frx":0E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBilance.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBilance.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBilance.frx":1DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBilance.frx":2336
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBilance.frx":287A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBilance.frx":2DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBilance.frx":343A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   84
      Top             =   8370
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   1111
      ButtonWidth     =   1296
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Caption         =   "Aggiunte"
            Key             =   "G"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Turno"
            Key             =   "T"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stampa"
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "phRight"
            Style           =   4
            Object.Width           =   500
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esci"
            Key             =   "X"
            Object.ToolTipText     =   "Esci"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cryStampa 
      Left            =   7560
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Parte"
      Height          =   255
      Left            =   5280
      TabIndex        =   115
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Par"
      Height          =   255
      Left            =   5760
      TabIndex        =   113
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   120
      Picture         =   "frmBilance.frx":35D6
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Materiale in produzione"
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
      Left            =   1800
      TabIndex        =   99
      Top             =   -120
      Width           =   5190
   End
   Begin VB.Label Label12 
      Caption         =   "MP"
      Height          =   255
      Left            =   480
      TabIndex        =   98
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "Descrizione"
      Height          =   255
      Left            =   1800
      TabIndex        =   97
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Qta"
      Height          =   255
      Left            =   3960
      TabIndex        =   96
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "Lotto"
      Height          =   255
      Left            =   7920
      TabIndex        =   95
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Mescola"
      Height          =   255
      Left            =   240
      TabIndex        =   94
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   "ODL"
      Height          =   255
      Left            =   240
      TabIndex        =   93
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "ID"
      Height          =   255
      Left            =   240
      TabIndex        =   92
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label20 
      Caption         =   "Data"
      Height          =   255
      Left            =   6360
      TabIndex        =   91
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label22 
      Caption         =   "Qta in Kg"
      Height          =   255
      Left            =   240
      TabIndex        =   90
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "SE"
      Height          =   255
      Left            =   8400
      TabIndex        =   89
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Molt"
      Height          =   255
      Left            =   2760
      TabIndex        =   88
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Ciclo"
      Height          =   255
      Left            =   4800
      TabIndex        =   87
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Cicli"
      Height          =   255
      Left            =   4200
      TabIndex        =   86
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Tot"
      Height          =   255
      Left            =   6720
      TabIndex        =   85
      Top             =   2640
      Width           =   255
   End
   Begin VB.Shape shpRisposta 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   5700
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2630
      Width           =   9255
   End
   Begin VB.Shape Shape2 
      Height          =   1935
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   560
      Width           =   9255
   End
End
Attribute VB_Name = "frmBilance"
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
   cAggiunte = 11
   cTurno = 13
   cStampa = 15
   cEsci = 17
End Enum
Private Enum Status
   cNew = 0
   cMod = 1
   cExists = 2
   cNotExists = 3
End Enum
Dim rsComponente As Recordset, rsN As Recordset
Dim fS As Boolean, dTot As Double, dC As Double, dF As Double
Dim fSalva As Boolean, dP As Double, sSE As String, dS As Double
Dim fConto As Boolean, fRicalcola
Public Property Let Status(vValue As Integer)
   If vValue >= cNew And vValue <= cNotExists Then
      mStatus = vValue
   End If
   With tbCommand
    If sFlag = "ATC" Or sFlag = "BIL" Then ' controllo che non possano modificare i dati dal MKG
      Select Case mStatus
         Case cNew, cMod
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = True
            .Buttons(cAnnulla).Enabled = True
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cAggiunte).Enabled = False
            .Buttons(cTurno).Enabled = False
            .Buttons(cEsci).Enabled = False
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = True
            .Buttons(cChiudi).Enabled = Not (sStato = "C")
            .Buttons(cStampa).Enabled = True
            .Buttons(cAggiunte).Enabled = Not (sStato = "C")
            .Buttons(cTurno).Enabled = True
            .Buttons(cEsci).Enabled = True
            fS = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = True
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cAggiunte).Enabled = False
            .Buttons(cTurno).Enabled = False
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
            .Buttons(cAggiunte).Enabled = True
            .Buttons(cTurno).Enabled = True
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cAggiunte).Enabled = False
            .Buttons(cTurno).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cAggiunte).Enabled = False
            .Buttons(cTurno).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
      End Select
    End If
    
    
    
   End With
End Property
Public Property Get Status() As Integer
   Status = mStatus
End Property

Public Property Let Key(vValue As String)
    mKey = Trim$(vValue)
    CercaID
    CercaComponenti
    
    NMescole
        
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

Private Sub cmdCalcola_Click()
Dim i As Integer
Dim sqlc As String, rsT As Recordset, dFine As Double
dC = 0
dF = 0
dP = 0
dFine = 0
fRicalcola = False
'If Not IsNull(rsN.RecordCount) And rsN.RecordCount <> "" Then
'        dFine = rsN.RecordCount
'Else
    
    sqlc = "SELECT MP.SeId, Count(MP.MP) AS CMP " & _
            "FROM SE INNER JOIN MP ON SE.SeId = MP.SeId " & _
            "GROUP BY MP.SeId " & _
            "HAVING MP.SeId=" & txtSeId.Text
    Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
        If Not rsT.EOF Then
            dFine = rsT("CMP")
        Else
            dFine = rsN.RecordCount
        End If
    rsT.Close
    Set rsT = Nothing
'End If

If Not IsNull(txtCicli.Text) And txtCicli.Text <> "" Then
        
    For i = 0 To dFine - 1
        txtFatt(i).Text = txtMPP(i).Text * txtF.Text
        txtFatt(i).Visible = True
        dF = dF + txtFatt(i).Text
        If Not IsNull(txtParte.Text) And txtParte.Text <> "" Then
            txtPC(i).Text = txtMPP(i).Text * txtParte.Text
            txtPC(i).Visible = True
            dP = dP + txtPC(i)
            txtFT(i).Text = (txtFatt(i).Text * txtCicli.Text) + txtPC(i).Text
            txtFT(i).Visible = True
            dC = dC + txtFT(i).Text
        Else
            txtFT(i).Text = txtFatt(i).Text * txtCicli.Text
            txtFT(i).Visible = True
            dC = dC + txtFT(i).Text
       
        End If
      
    Next i

Else
    For i = 0 To dFine - 1
        txtFatt(i).Text = txtMPP(i).Text * txtF
        txtFatt(i).Visible = True
        dF = dF + txtFatt(i).Text
        
    Next i
End If
    txtSF.Text = dF
    txtSFT.Text = dC
    txtSPC.Text = dP
    fConto = True
    fRicalcola = True
End Sub
Private Sub SalvaMP()
Dim sqlc As String, rsMp As Recordset, i As Integer, dTot As Double
Dim rsT As Recordset
dTot = 0
    sqlc = "SELECT MP.SeId, MP.MP, MP.Lotto, MP.Qta, " & _
                "MP.QtaC, MP.QtaT, MP.QtaP, SE.Ciclo, SE.Fattore, SE.Parte " & _
            "FROM SE INNER JOIN MP ON SE.SeId = MP.SeId " & _
            "WHERE MP.SeId=" & txtSeId.Text

    Set rsMp = db.OpenRecordset(sqlc, dbOpenDynaset)
    
If fNuovoRecord Then
'    Do While Not rsComponente.EOF
    For i = 0 To rsN.RecordCount - 1
        
        rsMp.AddNew
'    rsComponente.AddNew
        rsMp("SeId") = txtSeId.Text
        rsMp("MP") = txtMP(i).Text
        rsMp("Qta") = txtMPP(i).Text
        rsMp("QtaC") = txtFatt(i).Text
        rsMp("QtaT") = txtFT(i).Text
        rsMp("QtaP") = txtPC(i).Text
        If Not IsNull(txtLotto(i).Text) And txtLotto(i).Text <> "" Then
            rsMp("Lotto") = txtLotto(i).Text
        Else
            rsMp("Lotto") = "0"
        End If
        rsMp("Ciclo") = txtCicli.Text
        rsMp("Fattore") = txtF.Text
        If Not IsNull(txtParte.Text) And txtParte.Text <> "" Then
            rsMp("Parte") = txtParte.Text
        Else
            rsMp("Parte") = "0"
        End If
        
        rsMp.Update
'        rsMp.MoveNext
'        rsN.MoveNext
    Next i

Else
    rsMp.Edit
'Controllare per l'inserimento delle qta
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    For i = 0 To rsComponente.RecordCount - 1
        rsMp.Edit
            rsMp("MP") = txtMP(i).Text
            rsMp("Qta") = txtMPP(i).Text
            rsMp("QtaC") = txtFatt(i).Text
            rsMp("QtaT") = txtFT(i).Text
            rsMp("QtaP") = txtPC(i).Text
            rsMp("Ciclo") = txtCicli.Text
            rsMp("Fattore") = txtF.Text
            rsMp("Parte") = txtParte.Text

        If Not IsNull(txtLotto(i).Text) And txtLotto(i) <> "" Then
            rsMp("Lotto") = txtLotto(i).Text
        Else
            rsMp("Lotto") = "0"
        End If
            dTot = dTot + rsMp("QtaT")
        rsMp.Update
        rsMp.MoveNext
    Next i
'    Loop

End If
   If fRicalcola Then
    sqlc = "Select SeId, Qta " & _
            "from SE " & _
            "where SeId=" & txtSeId.Text
    Set rsT = db.OpenRecordset(sqlc, dbOpenDynaset)
        If Not rsT.EOF Then
            rsT.Edit
                rsT("Qta") = dTot
            rsT.Update
        End If
   End If
   
   
   fNuovoRecord = False
   CercaComponenti
'     disabilito i campi
'   Campi False
End Sub

Private Sub Form_Activate()
    NMescole
End Sub


Private Sub Form_Load()
   Dim btn As Button
  fS = False
   Campi False
   For Each btn In tbCommand.Buttons
     If btn.Key <> "X" Then
        btn.Enabled = False
     End If
   Next
   CampiD False
   fRicalcola = False
   
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
Private Sub CercaIDN()
Dim sqlc As String, i As Integer
dTot = 0
sSE = ValID(Me.Key)
    
    sqlc = "SELECT Formule.ID, Formule.Mescola, Form.MP, Form.Qta, " & _
                "DescrizMesc.Descrizione " & _
            "FROM (Formule LEFT JOIN Form ON Formule.Mescola = " & _
                "Form.Mescola) LEFT JOIN DescrizMesc ON Form.MP " & _
                "= DescrizMesc.MP " & _
            "GROUP BY Formule.ID, Formule.Mescola, Form.MP, " & _
                "Form.Qta, DescrizMesc.Descrizione " & _
            "HAVING Formule.ID='" & sSE & "'"


Set rsN = db.OpenRecordset(sqlc, dbOpenSnapshot)

            
            For i = 0 To rsN.RecordCount - 1
            If Not IsNull(rsN("MP")) Then
                txtMP(i).Visible = True
                txtMP(i).Text = rsN("MP")
            Else
                txtMP(i) = ""
            End If
            If Not IsNull(rsN("Descrizione")) Then
                txtMPD(i).Visible = True
                txtMPD(i).Text = rsN("Descrizione")
            Else
                txtMPD(i).Text = ""
            End If
            If Not IsNull(rsN("Qta")) Then
                txtMPP(i).Visible = True
                txtMPP(i).Text = Trim$(rsN("Qta"))
                dTot = dTot + txtMPP(i)
            Else
                txtMPP(i).Text = ""
            End If
'            If Not IsNull(rsN("Lotto")) Then
                txtLotto(i).Visible = True
 '               txtLotto(i).Text = rsN("Lotto")
  '          Else
                txtLotto(i).Text = ""
  '          End If
                
        rsN.MoveNext
        Next i
     
   txtTot.Text = dTot

End Sub
Private Sub CercaComponenti()
Dim sqlc As String, i As Integer, dTot As Double
Dim rsT As Recordset, dSe As Double
  sSE = ValID(Me.Key)
          
 If fNuovoRecord Then
    sqlc = "SELECT DescrizMesc.Descrizione, Form.MP, Form.Qta, " & _
                "Formule.ID, '' AS Lotto " & _
            "FROM (Form INNER JOIN Formule ON Form.Mescola " & _
                "= Formule.Mescola) LEFT JOIN DescrizMesc ON " & _
                "Form.MP = DescrizMesc.MP " & _
            "WHERE Formule.ID='" & sSE & "'"
    
Else
    sqlc = "SELECT Formule.ID, Formule.SeID " & _
            "From Formule " & _
            "Where Formule.ID = '" & sSE & "'"
    Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
        If Not rsT.EOF Then
            dSe = rsT("SeId")
        End If
    rsT.Close
    Set rsT = Nothing
    
    sqlc = "SELECT MP.MP, DescrizMesc.Descrizione, MP.Qta, " & _
                "MP.QtaC, MP.QtaT, MP.QtaP, MP.Lotto, SE.Ciclo, SE.Fattore, SE.Parte, SE.Qta " & _
            "FROM ((Formule INNER JOIN MP ON Formule.SeID " & _
                "= MP.SeId) LEFT JOIN DescrizMesc ON MP.MP = " & _
                "DescrizMesc.MP) INNER JOIN SE ON Formule.SeID = SE.SeId " & _
            "GROUP BY MP.MP, DescrizMesc.Descrizione, MP.Qta, " & _
                "MP.QtaC, MP.QtaT, MP.QtaP, MP.Lotto, MP.SeId, " & _
                "SE.Ciclo, SE.Fattore, SE.Parte, SE.Qta " & _
            "HAVING MP.SeId=" & dSe

End If

    
    Set rsComponente = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsComponente.EOF Then
        rsComponente.MoveLast
        
        i = rsComponente.RecordCount
        rsComponente.MoveFirst
        
        FoundID
        Me.Status = cExists
    Else
        Me.Status = cNotExists
    End If
End Sub
Private Sub FoundID()
Dim i As Integer, dTot As Double, dDiff As Double, fConta As Boolean
Dim iRes As Integer
dTot = 0
dC = 0
dF = 0
dP = 0
dS = 0
fConta = False
        For i = 0 To rsComponente.RecordCount - 1
            If Not IsNull(rsComponente("MP")) Then
                txtMP(i).Visible = True
                txtMP(i).Text = rsComponente("MP")
            Else
                txtMP(i) = ""
            End If
            If Not IsNull(rsComponente("Descrizione")) Then
                txtMPD(i).Visible = True
                txtMPD(i).Text = rsComponente("Descrizione")
            Else
                txtMPD(i).Text = ""
            End If
            'Controllo se già memorizzato
            If Not fNuovoRecord Then
'            If Not IsNull(rsComponente("Qta")) Then
                txtMPP(i).Text = rsComponente("MP.Qta")
            Else
                If Not IsNull(rsComponente("MP.Qta")) Then
                    txtMPP(i).Text = rsComponente("MP.Qta")
                Else
                    txtMPP(i).Text = ""
                End If
                
            End If
            dTot = dTot + txtMPP(i)
            txtMPP(i).Visible = True
            If Not fNuovoRecord Then
'            If Not IsNull(rsComponente("Qta")) Then
                txtFatt(i).Text = rsComponente("Qtac")
            Else
                If Not IsNull(rsComponente("QtaC")) Then
                    txtFatt(i).Text = rsComponente("QtaC")
                Else
                    txtFatt(i).Text = ""
                End If
                
            End If
            dF = dF + txtFatt(i)
            txtFatt(i).Visible = True
            If Not fNuovoRecord Then
'            If Not IsNull(rsComponente("Qta")) Then
                txtFT(i).Text = rsComponente("QtaT")
            Else
                If Not IsNull(rsComponente("QtaT")) Then
                    txtFT(i).Text = rsComponente("QtaT")
                Else
                    txtFT(i).Text = ""
                End If
                
            End If
            dC = dC + txtFT(i)
            txtFT(i).Visible = True
            If Not fNuovoRecord Then
'            If Not IsNull(rsComponente("Qta")) Then
                txtPC(i).Text = rsComponente("QtaP")
            Else
                If Not IsNull(rsComponente("QtaP")) Then
                    txtPC(i).Text = rsComponente("QtaP")
                Else
                    txtPC(i).Text = ""
                End If
                
            End If
            If Not IsNull(txtPC(i).Text) And txtPC(i).Text <> "Text1" Then
                dP = dP + txtPC(i)
                txtPC(i).Visible = True
            End If
            
            If Not IsNull(rsComponente("Lotto")) Then
                txtLotto(i).Visible = True
                txtLotto(i).Text = rsComponente("Lotto")
            Else
                txtLotto(i).Text = ""
            End If
            txtF.Text = rsComponente("Fattore")
            txtCicli.Text = rsComponente("Ciclo")
            txtParte.Text = rsComponente("Parte")
            If Not IsNull(rsComponente("Se.Qta")) And rsComponente("Se.Qta") <> "" Then
                If rsComponente("Se.Qta") > 0 Then
                   fConta = True
                    dS = rsComponente("SE.Qta")
                Else
                    fConta = False
                End If
            Else
                fConta = False
            End If
            
        rsComponente.MoveNext
        Next i
     
    txtTot.Text = dTot
    txtSF.Text = dF
    txtSFT.Text = dC
    txtSPC.Text = dP
'Da inserire
'if dc>rscomponente("SE.Qta") then
'   msgbox"Dovete modificare la quantità di un elemento di KG " & differenza tra dc e se.qta
'inserire obbligatoriamente
'   endif
    dDiff = Format((dC - dS), "0")
    If fConta Then
        If dDiff <> 0 Then
'            MsgBox "dovete modificare la quantità inserita di un elemento " & vbCrLf & _
'                    "di " & dDiff & "kg", vbExclamation
'            fConto = False
            iRes = MsgBox("Dovreste modificare la quantità inserita di un elemento " & vbCrLf & _
                    "di " & dDiff & "kg. Volete farlo?", _
                    vbYesNo + vbDefaultButton1 + vbQuestion)
            If iRes = vbNo Then
                fConto = True
            Else
                fConto = False
            End If
        Else
            fConto = True
        End If
    End If

End Sub
Private Sub CercaID()
Dim sqlc As String, rsId As Recordset, i As Integer, rsT As Recordset
Dim dQ As Double

'   s = Me.Key
'   iPos = InStr(s, ",")
'   If iPos <> 0 Then
'      s = Left$(s, iPos - 1)
'   End If
sSE = ValID(Me.Key)
sqlc = "SELECT Mescola, Data, OdL, ID, CodArt, Qta, Caricata, SeId " & _
        "From Formule " & _
        "WHERE ID='" & sSE & "'"
Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsT.EOF Then
        sSE = rsT("SeId")
    End If
'sqlc = "SELECT Formule.Mescola, Formule.Data, Formule.OdL, " & _
            "Formule.ID, Formule.CodArt, Formule.Qta, Formule.Caricata, Formule.SeID " & _
        "From Formule " & _
        "WHERE Formule.SeID=" & sSe
sqlc = "SELECT Formule.Mescola, DescrizMesc.Descrizione, " & _
            "Formule.Data, Formule.OdL, Formule.ID, " & _
            "Formule.CodArt, Formule.Qta, Formule.Caricata, Formule.SeID " & _
        "FROM Formule LEFT JOIN DescrizMesc ON Formule.Mescola " & _
            "= DescrizMesc.MP " & _
        "WHERE Formule.SeID=" & sSE
Set rsId = db.OpenRecordset(sqlc, dbOpenDynaset)
        
'Inizializziamo
txtId.Text = ""
txtODL.Text = ""
txtQta.Text = ""
dQ = 0
        
        Do While Not rsId.EOF
                        
            txtMescola.Text = rsId("Mescola")
            
            If Not IsNull(rsId("Descrizione")) Then
                txtDM.Text = rsId("Descrizione")
            End If
            If Not IsNull(txtId.Text) And txtId.Text <> "" Then
                txtId.Text = txtId.Text & ", " & rsId("ID")
            Else
                txtId.Text = rsId("ID")
            End If
            If Not IsNull(txtODL.Text) And txtODL.Text <> "" Then
                txtODL.Text = txtODL.Text & ", " & rsId("ODL")
            Else
                txtODL.Text = rsId("ODL")
            End If
            If Not IsNull(txtQta.Text) And txtQta.Text <> "" Then
                dQ = dQ + rsId("Qta")
            Else
                dQ = rsId("Qta")
            End If
            txtQta.Text = dQ
            
            txtSeId.Text = rsId("SeId")
            txtData.Text = Date
            
        rsId.MoveNext
        Loop

End Sub

Private Sub Campi(Flag As Boolean)
    Dim i As Integer, dr As Double
    
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
'        txtCMP(i).Visible = Flag
'        txtCMPQ(i).Visible = Flag
        txtFatt(i).Visible = Flag
        txtFT(i).Visible = Flag
        txtPC(i).Visible = Flag
    Next i
    
    
    
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
            CercaIDN
            Me.Status = cNew
            CampiD True
        Case "S"
            Controllo
            If fSalva Then
                SalvaMP
                Me.Status = cExists
                '        Campi False
                CampiD False
            End If
            
        Case "A"
            If fNuovoRecord Then
'                CancellaCampi
                Me.Status = cNotExists
            Else
                CercaComponenti
                Me.Status = cExists
            End If
            '         Campi False
            CampiD False
        Case "M"
            fNuovoRecord = False
            Campi True
            Me.Status = cMod
            Controtipo
            CampiD True
        Case "C"
            iRes = MsgBox("Vuoi chiudere questa produzione?", _
            vbYesNo + vbDefaultButton2 + vbQuestion)
            If iRes = vbYes Then
                If fConto Then
                    sqlc = "Update Formule set Caricata = 'F' " & _
                        "Where SeId = " & txtSeId.Text & " and " & _
                        "Caricata = 'L'"
                    db.Execute (sqlc)
                    sqlc = "Update SE set Qta = '" & txtSFT.Text & "' " & _
                            "where SeId = " & txtSeId.Text
                    db.Execute (sqlc)
'                    sStato = "C"
                    Me.Status = cExists
                    frmTV.ChangedTV = True
                    Unload Me
                Else
                    MsgBox "Le quantità impostate non sono valide", vbExclamation
                    
                End If
            End If
        Case "P"
            fNuovoRecord = False
'            Rete
'            Me.PrintForm
            Stampa
            Me.Status = cExists
        Case "G"
'            Load frmPass
'            frmPass.Key = txtSeId.Text
'            frmPass.Show
            Load frmAggiunte
            frmAggiunte.Key = txtSeId.Text
            frmAggiunte.Show
            Unload Me
        Case "T"
            Load frmTurno
            frmTurno.Key = txtSeId.Text
            frmTurno.Show
            NMescole
        
    End Select
End Sub

Private Sub Controtipo()
'Dim sqlc As String, rsCon As Recordset
'Dim sWhere As String, i As Integer
'
'    For i = 0 To rsComponente.RecordCount - 1
'        sWhere = txtMP(i).Text
'        sqlc = "SELECT MP, Controtipo " & _
'                "From Controtipo " & _
'                "WHERE MP='" & sWhere & "'"
'        Set rsCon = db.OpenRecordset(sqlc, dbOpenDynaset)
'            If Not rsCon.EOF Then
'                txtCMP(i).Text = rsCon("Controtipo")
'                txtCMPQ(i).Visible = True
'     '           txtCMPQ(i).Text = ""
'            Else
'                txtCMP(i).Text = ""
'                txtCMP(i).Visible = False
'                txtCMPQ(i).Visible = False
'            End If
'                txtCMPQ(i).Text = ""
'
'    Next i
        
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
'        txtCMP(i).Enabled = Flag
'        txtCMPQ(i).Enabled = Flag
        txtFatt(i).Enabled = Flag
        txtFT(i).Enabled = Flag
        txtPC(i).Enabled = Flag
    Next i

End Sub

Private Sub Controllo()

If Not IsNull(txtF.Text) And txtF.Text <> "" Then
    If Not IsNull(txtCicli.Text) And txtCicli.Text <> "" Then
        ControlloQta
    Else
        fSalva = False
        MsgBox "Prima di salvare bisogna inserire i dati relativi ai cicli", vbExclamation
    End If
Else
    fSalva = False
    MsgBox "Prima di salvare bisogna inserire i dati relativi ai " & _
            "fattori di moltiplicazione", vbExclamation
End If


End Sub
Private Sub ControlloQta()
Dim iRes As Integer, sqlc As String
Dim dQta As Double, dSomma As Double, dEcc As Double

dQta = txtQta.Text * 1.1
If Not IsNull(txtSFT.Text) And txtSFT.Text <> "" Then
    dSomma = txtSFT.Text
Else
    fSalva = False
    MsgBox "Inserire i cicli ed il fattore di moltiplicazione", vbCritical
    Exit Sub
End If
dEcc = txtQta.Text * 1.2
    If dSomma < dQta Then
        iRes = MsgBox("La quantità in preparazione è inferiore all'ordine " & _
                    " del cliente più un 10%. Volete continuare?", _
               vbYesNo + vbDefaultButton2 + vbQuestion)
        If iRes = vbYes Then
            fSalva = True
        Else
            fSalva = False
        End If

    Else
        If dSomma > dEcc Then
            iRes = MsgBox("La quantità in preparazione è superiore all'ordine " & _
                        " del cliente più un 20%. Volete continuare?", _
                   vbYesNo + vbDefaultButton2 + vbQuestion)
            If iRes = vbYes Then
                fSalva = True
            Else
                fSalva = False
            End If
        Else
            fSalva = True
        End If
        
    End If
    


End Sub

Private Sub Stampa()
'   Dim dDataInizio As Date, dDataFine As Date,
  Dim sSelFormula As String
'   Dim sTitolo As String
   Dim sTmp As String



   sTmp = txtSeId.Text
'   If mskDataA.Text = "__/____" Then
'      mskDataA.Text = mskDataDA.Text
'   End If
'   dDataFine = DateAdd("d", -1, DateAdd("m", 1, "01/" & mskDataA.Text))
'   dDataInizio = DateSerial(Val(Mid$(mskDataDA.Text, 4, 4)), _
'                            Val(Mid$(mskDataDA.Text, 1, 2)), 1)
'   If dDataFine <= dDataInizio Then
'      MsgBox "Date Errate !!"
'      Exit Sub
'   End If
''
''  Preparo per il titolo
''
'   If mskDataDA.Text = mskDataA.Text Then
'      sTitolo = "Indice di Qualità per il mese di: " & _
'                  Format$(dDataInizio, "mmmm yyyy")
'   Else
'      If Year(dDataInizio) = Year(dDataFine) Then
'         sTitolo = "Dettaglio Reclami per il periodo: " & _
'                     Format$(dDataInizio, "mmmm") & " - " & _
'                     Format$(dDataFine, "mmmm yyyy")
'      Else
'         sTitolo = "Dettaglio Reclami per il periodo: " & _
'                     Format$(dDataInizio, "mmmm yyyy") & " - " & _
'                     Format$(dDataFine, "mmmm yyyy")
'      End If
'   End If
   
   With cryStampa
      On Error Resume Next
      .Destination = crptToWindow
      .ReportFileName = App.Path & "\bilance.rpt"
'      .Formulas(0) = "Titolo=" & Chr$(34) & sTitolo & Chr$(34)
      .WindowState = crptNormal
      sSelFormula = "{qBilance.SeId}=" & sTmp

      .SelectionFormula = sSelFormula
      .Action = 1
         If Err.Number <> 0 Then
          MsgBox "Errore nel produrre il report " & vbCrLf & _
              Err.Number & " - " & Err.Description
          Err.Clear
      End If
   End With

End Sub

Private Sub NMescole()
Dim sqlc As String, rsMe As Recordset
Dim sMe As String, fMe As Boolean

fMe = False

sqlc = "Select SeId, Mescole " & _
        "From Turno " & _
        "Where SeId=" & txtSeId
        
Set rsMe = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsMe.EOF Then
        Do While Not rsMe.EOF
            If fMe Then
                sMe = sMe + Val(rsMe("Mescole"))
            Else
                sMe = rsMe("Mescole")
                fMe = True
            End If
            rsMe.MoveNext
        Loop
        txtMescole.Text = "Mescole " & sMe & " di " & txtCicli.Text
        txtMescole.Visible = True
        txtMescole.Locked = True
    Else
        txtMescole.Visible = False
    End If


End Sub

