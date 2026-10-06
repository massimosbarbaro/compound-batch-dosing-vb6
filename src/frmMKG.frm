VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMKG 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Marketing"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7800
   Icon            =   "frmMKG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNote 
      Height          =   1005
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Top             =   6720
      Width           =   4335
   End
   Begin VB.TextBox txtReclamo 
      Height          =   615
      Index           =   9
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   5400
      Width           =   5880
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   16
      Left            =   480
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtReclamo 
      Height          =   615
      Index           =   8
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Top             =   4680
      Width           =   5880
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   15
      Left            =   5985
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   4080
      Width           =   870
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   11
      Left            =   1200
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   10
      Left            =   1200
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   4
      Left            =   6240
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtReclamo 
      Height          =   645
      Index           =   3
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Text            =   "frmMKG.frx":0CCA
      Top             =   2160
      Width           =   5055
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   2
      Left            =   3120
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtReclamo 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtNota 
      Height          =   285
      Left            =   1080
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox txtAccr 
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtReso 
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Frame fraEsterno 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Esterno"
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
      Begin VB.TextBox txtReclamo 
         Height          =   285
         Index           =   12
         Left            =   360
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraIQ 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Indice di Qualità"
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
      Begin VB.TextBox txtReclamo 
         Height          =   285
         Index           =   13
         Left            =   360
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraSoluzione 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Soluzione"
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
      Begin VB.TextBox txtReclamo 
         Height          =   285
         Index           =   14
         Left            =   120
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdDettagli 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Dettagli"
      Height          =   615
      Left            =   6600
      MousePointer    =   1  'Arrow
      Picture         =   "frmMKG.frx":0CD0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   855
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   21
      Top             =   7965
      Width           =   7800
      _ExtentX        =   13758
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
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   6720
      Top             =   0
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
            Picture         =   "frmMKG.frx":1202
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMKG.frx":135E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMKG.frx":18A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMKG.frx":1DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMKG.frx":232A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMKG.frx":286E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMKG.frx":2DB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note "
      Height          =   195
      Left            =   4920
      TabIndex        =   42
      Top             =   6480
      Width           =   390
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Analisi dei reclami"
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
      TabIndex        =   41
      Top             =   480
      Width           =   4035
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nota di credito"
      Height          =   435
      Left            =   240
      TabIndex        =   39
      Top             =   6960
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accredito £"
      Height          =   195
      Left            =   240
      TabIndex        =   38
      Top             =   7485
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reso KG"
      Height          =   195
      Left            =   240
      TabIndex        =   37
      Top             =   6525
      Width           =   645
   End
   Begin VB.Label lblDifetto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Difetto"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   3840
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Famiglia"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   4200
      Width           =   570
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Azioni Correttive"
      Height          =   435
      Left            =   360
      TabIndex        =   15
      Top             =   5280
      Width           =   780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Risposta"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Risposta"
      Height          =   195
      Left            =   2040
      TabIndex        =   12
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Risposta N"
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   3120
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qta Accettata"
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   3840
      Width           =   990
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ODL"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   330
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Reclamo"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qta Reclamata"
      Height          =   195
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label lblReclamoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Relcamo Numero"
      Height          =   195
      Left            =   4800
      TabIndex        =   3
      Top             =   1560
      Width           =   1230
   End
   Begin VB.Shape Shape2 
      Height          =   1575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Motivazioni"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venditore"
      Height          =   195
      Left            =   6600
      TabIndex        =   1
      Top             =   1560
      Width           =   675
   End
   Begin VB.Image imgLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmMKG.frx":32F6
      Top             =   0
      Width           =   450
   End
   Begin VB.Shape shpRisposta 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   3255
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   7455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   7455
   End
End
Attribute VB_Name = "frmMKG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMKG As Recordset, sStatoReclamo As String, fNuovoRecord As Boolean
Private mKey As String
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
Public Property Let Status(vValue As Integer)
   If vValue >= cNew And vValue <= cNotExists Then
      mStatus = vValue
   End If
   With tbCommand
    If sFlag = "MKG" Then 'Controllo che non possano modificare i dati da ATC
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
            .Buttons(cChiudi).Enabled = (sStatoReclamo = "V")
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = True
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
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


Public Property Let Key(ByVal vValue As String)
   mKey = Trim$(vValue)
   CaricoCampi
   CercaReclamoMKG
 
 
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub CaricoCampi()
 Dim sqlc As String, rsReclamo As Recordset, i As Integer
    sqlc = "SELECT ATC.ODL, RV.DataReclamo, RV.QtaReclamata, " & _
            "RV.CausaReclamo, RV.Venditore, RV.NumeroReclamo, " & _
            "ATC.NumeroRisposta, ATC.DataRisposta, ATC.Risposta, " & _
            "ATC.AzioniCorrettive, Famiglia.DescrizioneFamiglia, " & _
            "Difetto.DescrizioneDifetto, IIf([Esterno]=0,'No','Sì') AS Ext, " & _
            "IIf([IndiceQualita]=0,'No','Si') AS IQ, " & _
            "IIf([Soluzione]='0','Accettato',(IIf([Soluzione]='1','Respinto','Segnalazione'))) " & _
            "AS Sol, ATC.QtaAccettata, ATC.StatoReclamo " & _
            "FROM Famiglia RIGHT JOIN (Difetto RIGHT JOIN (ATC INNER JOIN ReclamoVenditore AS RV " & _
            "ON ATC.ODL = RV.ODL) ON Difetto.CodDifetto = ATC.CodDifetto) " & _
            "ON Famiglia.CodFamiglia = ATC.CodFamiglia " & _
            "where ATC.ODL = '" & Me.Key & "'"
    Set rsReclamo = db.OpenRecordset(sqlc, dbOpenDynaset)
   If Not rsReclamo.EOF Then
     If sFlag = "MKG" Then 'controlo chi apre il DB
      If rsReclamo("StatoReclamo") = "N" Then
         rsReclamo.Edit
         rsReclamo("StatoReclamo") = "V"
         rsReclamo.Update
         frmTV.ChangedTV = True
      End If
     End If
      For i = 0 To rsReclamo.Fields.Count - 1
         If Not IsNull(rsReclamo(i)) Then
               txtReclamo(i).Text = rsReclamo(i)
            Else
                txtReclamo(i) = ""
         End If
      Next i
      sStatoReclamo = rsReclamo("StatoReclamo")
   Else
      Err.Raise 13, , "Errore nel data source"
   End If
    
    rsReclamo.Close
    Set rsReclamo = Nothing
End Sub
Private Sub CercaReclamoMKG()
 Dim sqlc As String, i As Integer
   
    sqlc = "SELECT ODL, Accredito, NotaCredito, Reso, Note " & _
           "FROM MKG where ODL='" & Me.Key & "'"
    Set rsMKG = db.OpenRecordset(sqlc, dbOpenDynaset)
    
   If Not rsMKG.EOF Then
      FoundODL
      Me.Status = cExists
    Else
      Me.Status = cNotExists
    End If
End Sub
Private Sub FoundODL()
    If Not IsNull(rsMKG("NotaCredito")) Then
        txtNota = rsMKG("NotaCredito")
    Else
        txtNota = ""
    End If
    If Not IsNull(rsMKG("Accredito")) Then
        txtAccr = rsMKG("Accredito")
    Else
        txtAccr = ""
    End If
    If Not IsNull(rsMKG("Reso")) Then
        txtReso = rsMKG("Reso")
    Else
        txtReso = ""
    End If
    If Not IsNull(rsMKG("Note")) Then
      txtNote.Text = rsMKG("Note")
   Else
      txtNote = ""
   End If
   
End Sub

Private Sub cmdDettagli_Click()
    Dim lDiff As Long
    Load frmMFG
    With frmMFG
        .Top = Me.Top
        .Left = Me.Left + Me.Width
'   Controllo di non superare le dimensioni dello schermo
'       altrimenti la sposto indietro
        lDiff = Screen.Width * 0.98 - (.Left + .Width)
        If lDiff < 0 Then
            .Left = .Left + lDiff
        End If
        .Key = Me.Key
        .Show
    End With

End Sub

Private Sub cmdEsci_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim btn As Button
   Dim i As Integer
   For i = 0 To Me.txtReclamo.Count - 1
        txtReclamo(i).Enabled = True
        txtReclamo(i).Locked = True
   Next i
   
   
   Campi False
   For Each btn In tbCommand.Buttons
     If btn.Key <> "X" Then
        btn.Enabled = False
     End If
   Next

End Sub
Private Sub Campi(Flag As Boolean)
    txtAccr.Enabled = Flag
    txtReso.Enabled = Flag
    txtNota.Enabled = Flag
    txtNote.Enabled = Flag
    txtAccr.Locked = Not Flag
    txtReso.Locked = Not Flag
    txtNota.Locked = Not Flag
    txtNote.Locked = Not Flag
End Sub
Private Sub CancellaCampi()
   txtAccr.Text = ""
   txtReso.Text = ""
   txtNota.Text = ""
   txtNote.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   rsMKG.Close
   Set rsMKG = Nothing
End Sub

Private Sub tbCommand_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim iRes As Integer, rsTmp As Recordset, sqlc As String
   Select Case Button.Key
      Case "X"
         Unload Me
      Case "N"
         fNuovoRecord = True
         Campi True
         CancellaCampi
         Me.Status = cNew
      Case "S"
         If CheckCampi Then
            SalvaRec
            Me.Status = cExists
         End If
      Case "A"
         If fNuovoRecord Then
            CancellaCampi
            Me.Status = cNotExists
         Else
            FoundODL
            Me.Status = cExists
         End If
         Campi False
      Case "M"
         fNuovoRecord = False
         Campi True
         Me.Status = cMod
      Case "C"
        iRes = MsgBox("Vuoi chiudere il corrente reclamo", _
               vbYesNo + vbDefaultButton2 + vbQuestion)
         If iRes = vbYes Then
            sqlc = "Update ATC set StatoReclamo = 'C' " & _
                     "Where ODL = '" & txtReclamo(0) & "' and " & _
                     "StatoReclamo = 'V'"
            db.Execute (sqlc)
            sStatoReclamo = "C"
            Me.Status = cExists
            frmTV.ChangedTV = True
         End If
      Case "P"
         fNuovoRecord = False
         Rete
         Me.Status = cExists
   End Select
End Sub
Private Sub SalvaRec()
   Dim scod As Variant
   If fNuovoRecord Then
      rsMKG.AddNew
   Else
      rsMKG.Edit
   End If
   rsMKG("ODL") = Me.Key
   If Not (IsNull(txtReso.Text) Or txtReso.Text = "") Then
      rsMKG("Reso") = txtReso.Text
   Else
       rsMKG("Reso") = Null
   End If
   If Not (IsNull(txtNote.Text) Or txtNote.Text = "") Then
      rsMKG("Note") = txtNote.Text
   Else
      rsMKG("Note") = Null
   End If
   
   If Not (IsNull(txtAccr.Text) Or txtAccr.Text = "") Then
        rsMKG("Accredito") = txtAccr.Text
   Else
         rsMKG("Accredito") = Null
   End If
   If Not (IsNull(txtNota.Text) Or txtNota.Text = "") Then
       rsMKG("NotaCredito") = txtNota.Text
   Else
       rsMKG("NotaCredito") = Null
   End If
   
   rsMKG.Update
   CercaReclamoMKG
'     disabilito i campi
   Campi False
   fNuovoRecord = False
End Sub
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
        .Key = txtReclamo(6).Text
        .Show vbModal
    End With
End Sub
Private Function CheckCampi() As Boolean
   Dim fExit As Boolean
   fExit = True
   If IsNull(txtReso.Text) Or txtReso.Text = "" Then
         MsgBox "Inserire la quantità di materiale resa"
         txtReso.SetFocus
         fExit = False
   End If
   If fExit Then
      If IsNull(txtAccr.Text) Or txtAccr.Text = "" _
         Or Not IsNumeric(txtAccr.Text) Then
         MsgBox "Inserire il valore accreditato in formato numerico"
         txtAccr.SetFocus
         fExit = False
      End If
   End If
   If fExit Then
      If IsNull(txtNota.Text) Or txtNota.Text = "" Then
         MsgBox "Inserire la nota di credito"
         txtNota.SetFocus
         fExit = False
      End If
   End If
  
   
   CheckCampi = fExit
End Function

