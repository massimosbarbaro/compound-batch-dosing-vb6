VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReclamiVenditori 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Venditore"
   ClientHeight    =   5070
   ClientLeft      =   3075
   ClientTop       =   3540
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   9330
   Begin MSMask.MaskEdBox mskDataReclamo 
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Top             =   2760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.ListBox lstODL 
      Height          =   2985
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtCausaReclamo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3240
      Width           =   5535
   End
   Begin VB.TextBox txtQtaReclamata 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtNumeroReclamo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2760
      Width           =   735
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilToolBarV"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuovo"
            Key             =   "N"
            Object.ToolTipText     =   "Nuovo Reclamo"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
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
            Object.Visible         =   0   'False
            Caption         =   "Chiudi"
            Key             =   "C"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "phRight"
            Style           =   4
            Object.Width           =   500
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esci"
            Key             =   "X"
            Object.ToolTipText     =   "Esci"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToolBarV 
      Left            =   8160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":06A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":1128
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReclamiVenditori.frx":166C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   7095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qta"
      Height          =   195
      Left            =   8280
      TabIndex        =   20
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prodotto"
      Height          =   195
      Left            =   6240
      TabIndex        =   19
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente "
      Height          =   195
      Left            =   3840
      TabIndex        =   18
      Top             =   1200
      Width           =   525
   End
   Begin VB.Label lblQta 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   8040
      TabIndex        =   17
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblProdotto 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   5400
      TabIndex        =   16
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblCliente 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblODL 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ODL"
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Reclamo"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qta Reclamata"
      Height          =   195
      Left            =   4080
      TabIndex        =   9
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label lblReclamoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Relcamo Numero"
      Height          =   195
      Left            =   5760
      TabIndex        =   8
      Top             =   2520
      Width           =   1230
   End
   Begin VB.Shape shpRisposta 
      Height          =   1695
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   7095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Motivazioni"
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   795
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venditore"
      Height          =   195
      Left            =   7560
      TabIndex        =   6
      Top             =   2520
      Width           =   675
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmReclamiVenditori.frx":1BB0
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reclami Venditori"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   4005
   End
End
Attribute VB_Name = "frmReclamiVenditori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LB_FINDSTRING = &H18F
Private sSearch As String
Private mKey As String, rsReclamoVenditore As Recordset
Private fNuovoRecord As Boolean, sStato As String
Private Enum CMDS
   cNuovo = 1
   cSalva = 3
   cAnnulla = 5
   cModifica = 7
   cChiudi = 9
   cEsci = 11
End Enum
Private Function CheckCampi() As Boolean
   Dim fExit As Boolean
   fExit = True
   If (IsNull(mskDataReclamo.Text) Or mskDataReclamo.Text = "__/__/__") _
         Or Not IsDate(mskDataReclamo.Text) Then
      MsgBox "Data Errata"
      mskDataReclamo.SetFocus
      fExit = False
   End If
   If fExit Then
      If IsNull(txtQtaReclamata.Text) Or txtQtaReclamata.Text = "" Then
         MsgBox "Inserire la quantità di materiale reclamata"
         txtQtaReclamata.SetFocus
         fExit = False
      End If
   End If
   If fExit Then
      If IsNull(txtCausaReclamo.Text) Or txtCausaReclamo.Text = "" Then
         MsgBox "Inserire la causa del reclamo"
         txtCausaReclamo.SetFocus
         fExit = False
      End If
   End If
   If fExit Then
      If IsNull(txtVenditore.Text) Or txtVenditore.Text = "" Then
         MsgBox "Inserire il venditore che ha spedito il reclamo"
         txtVenditore.SetFocus
         fExit = False
      End If
   End If
   If fExit Then
      If IsNull(txtNumeroReclamo.Text) Or txtNumeroReclamo.Text = "" Then
         MsgBox "Inserire il numero del reclamo venditore"
         txtNumeroReclamo.SetFocus
         fExit = False
      End If
   End If
   
   CheckCampi = fExit
End Function

'Public Property Let Key(vValue As String)
'    mKey = Trim$(vValue)
'    Dim sqlc As String, i As Integer
'    CercaReclamoVenditore
'
'
'End Property
'Public Property Get Key() As String
'    Key = mKey
'End Property


'Private Sub cmdDettagli_Click()
'    Dim lDiff As Long
'    Load frmMFG
'    With frmMFG
'        .Top = Me.Top
'        .Left = Me.Left + Me.Width
'   Controllo di non superare le dimensioni dello schermo
'       altrimenti la sposto indietro
'        lDiff = Screen.Width * 0.98 - (.Left + .Width)
'        If lDiff < 0 Then
'            .Left = .Left + lDiff
'        End If
'        .Key = Me.Key
'        .Show vbModal
'    End With
'End Sub




Private Sub SalvaRec()
   Dim sCod As Variant
   If fNuovoRecord Then
      rsReclamoVenditore.AddNew
   Else
      rsReclamoVenditore.Edit
   End If
   
   rsReclamoVenditore("ODL") = lblODL.Caption
   rsReclamoVenditore("DataReclamo") = mskDataReclamo.Text
   rsReclamoVenditore("QtaReclamata") = txtQtaReclamata.Text
   rsReclamoVenditore("CausaReclamo") = txtCausaReclamo.Text
   rsReclamoVenditore("Venditore") = txtVenditore.Text
   rsReclamoVenditore("NumeroReclamo") = txtNumeroReclamo.Text
   rsReclamoVenditore("Stato") = "N"
   rsReclamoVenditore.Update
'   rsReclamoVenditore.MoveFirst
   CercaReclamoVenditore lblODL.Caption
'     disabilito i campi
   Campi False
   fNuovoRecord = False
End Sub



Private Sub Form_Load()
   Dim btn As Button
   Dim sqlc As String, rsODL As Recordset

   Me.MousePointer = vbArrowHourglass
   Me.Show
   Me.Refresh
   ShowWait "Attendere !!" & vbCrLf & vbCrLf & "Caricamento ODL in corso"
   sqlc = "Select *   from [InvioVenditore] order by 1"
   Set rsODL = db.OpenRecordset(sqlc, dbOpenSnapshot)
   Do While Not rsODL.EOF
      lstODL.AddItem rsODL(0)
      rsODL.MoveNext
   Loop
'   rsODL.Close
'   Set rsODL = Nothing
   lstODL.ListIndex = -1
   lstODL.SetFocus
   
   Campi False
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
   Me.MousePointer = vbNormal
   HideWait
End Sub

Private Sub Form_Resize()
   Dim b As Button, l As Long, idx As Integer
   If Me.WindowState <> vbMinimized Then
      l = shpRisposta.Left + shpRisposta.Width + Me.Width - Me.ScaleWidth
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
'   rsReclamoVenditore.Close
   Set rsReclamoVenditore = Nothing
End Sub

Private Function CercaReclamoVenditore(sODL As String) As Boolean
   Dim sqlc As String, i As Integer
   
   sqlc = "SELECT ODL, DataReclamo, QtaReclamata, " & _
          "CausaReclamo, Venditore, NumeroReclamo, Stato " & _
          "FROM ReclamoVenditore " & _
          "where ODL='" & Trim$(sODL) & "'"
   Set rsReclamoVenditore = db.OpenRecordset(sqlc, dbOpenDynaset)
   
   If Not rsReclamoVenditore.EOF Then
      CercaReclamoVenditore = True
   Else
      CercaReclamoVenditore = False
   End If
End Function
    
Sub pippo()
   FoundODL
   tbCommand.Buttons(cModifica).Enabled = True
   tbCommand.Buttons(cChiudi).Enabled = (sStato = "V")
   NuovoReclamo
End Sub
Private Sub FoundODL()
   If Not IsNull(rsReclamoVenditore("DataReclamo")) Then
      mskDataReclamo = rsReclamoVenditore("DataReclamo")
   Else
      mskDataReclamo = "__/__/__"
   End If
   If Not IsNull(rsReclamoVenditore("QtaReclamata")) Then
      txtQtaReclamata = rsReclamoVenditore("QtaReclamata")
   Else
      txtQtaReclamata = ""
   End If
   If Not IsNull(rsReclamoVenditore("CausaReclamo")) Then
      txtCausaReclamo = rsReclamoVenditore("CausaReclamo")
   Else
      txtCausaReclamo = ""
   End If
   If Not IsNull(rsReclamoVenditore("Venditore")) Then
      txtVenditore = rsReclamoVenditore("Venditore")
   Else
      txtVenditore = ""
   End If
   If Not IsNull(rsReclamoVenditore("NumeroReclamo")) Then
      txtNumeroReclamo = rsReclamoVenditore("NumeroReclamo")
   Else
      txtNumeroReclamo = ""
   End If
   tbCommand.Buttons(cNuovo).Enabled = False
   tbCommand.Buttons(cSalva).Enabled = False
   tbCommand.Buttons(cAnnulla).Enabled = False
   tbCommand.Buttons(cModifica).Enabled = True
End Sub
Private Sub Campi(Flag As Boolean)
   mskDataReclamo.Enabled = Flag
   txtQtaReclamata.Enabled = Flag
   txtCausaReclamo.Enabled = Flag
   txtVenditore.Enabled = Flag
   txtNumeroReclamo.Enabled = Flag
    
   txtQtaReclamata.Locked = Not Flag
   txtCausaReclamo.Locked = Not Flag
   txtVenditore.Locked = Not Flag
   txtNumeroReclamo.Locked = Not Flag
End Sub

Private Sub NuovoReclamo()
   tbCommand.Buttons(cNuovo).Enabled = True
   tbCommand.Buttons(cSalva).Enabled = False
   tbCommand.Buttons(cAnnulla).Enabled = False
   tbCommand.Buttons(cModifica).Enabled = False
End Sub
Private Sub CancellaCampi()
   mskDataReclamo.Text = "__/__/__"
   txtQtaReclamata.Text = ""
   txtCausaReclamo.Text = ""
   txtVenditore.Text = ""
   txtNumeroReclamo.Text = ""
End Sub


Private Sub lstODL_DblClick()
    Dim sTmp As String, rsTemp As Recordset, sqlc As String
   CancellaCampi
   sTmp = lstODL.List(lstODL.ListIndex)
   lblODL.Caption = sTmp
   sqlc = "Select *   from [InvioVenditore] where ODL='" & sTmp & "'"
   Set rsTemp = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsTemp.EOF Then
      lblCliente.Caption = rsTemp("Cliente")
      lblProdotto.Caption = rsTemp("Prodotto")
      lblQta.Caption = rsTemp("QtaSpedita")
   End If
   If CercaReclamoVenditore(sTmp) Then
      FoundODL
   Else
      NuovoReclamo
   End If
End Sub

Private Sub lstODL_KeyPress(KeyAscii As Integer)
'
'  Sarebbe anche meglio decidere quali tasti vanno ad accrescere la "Stringa"
'     (ad es. solo numeri ed il . od anche altri ??)
'
'  Inoltre potrebbe essere utile (per l'operatore) associare al tasto INVIO
'      l 'equivalente del DblClick sull'elemento
'
   If KeyAscii = vbKeyReturn Then
      lstODL_DblClick
      Exit Sub
   End If
   If KeyAscii = vbKeyBack Then
       If Len(sSearch) >= 1 Then
           sSearch = Left(sSearch, Len(sSearch) - 1)
       End If
   Else
      sSearch = sSearch & Chr$(KeyAscii)
   End If
   lstODL.ListIndex = SendMessage(lstODL.hwnd, LB_FINDSTRING, -1, ByVal CStr(sSearch))
   KeyAscii = 0
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
         With tbCommand
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = True
            .Buttons(cAnnulla).Enabled = True
            .Buttons(cModifica).Enabled = False
'            .Buttons(cChiudi).Enabled = False
            .Buttons(cEsci).Enabled = False
            lstODL.Enabled = False
         End With
      Case "S"
         If CheckCampi Then
            SalvaRec
            With tbCommand
               .Buttons(cNuovo).Enabled = False
               .Buttons(cSalva).Enabled = False
               .Buttons(cAnnulla).Enabled = False
               .Buttons(cModifica).Enabled = True
   '            .Buttons(cChiudi).Enabled = (sStato = "V")
               .Buttons(cEsci).Enabled = True
               lstODL.Enabled = True
            End With
         End If
      Case "A"
         If fNuovoRecord Then
            CancellaCampi
         Else
            FoundODL
         End If
         Campi False
         With tbCommand
            .Buttons(cNuovo).Enabled = fNuovoRecord    '  Dipende se arrivo da Mod o Nuovo
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = Not fNuovoRecord '  Dipende se arrivo da Mod o Nuovo
'            .Buttons(cChiudi).Enabled = (sStato = "V")
            .Buttons(cEsci).Enabled = True
            lstODL.Enabled = True
         End With
      Case "M"
         fNuovoRecord = False
         Campi True
         With tbCommand
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = True
            .Buttons(cAnnulla).Enabled = True
            .Buttons(cModifica).Enabled = False
'            .Buttons(cChiudi).Enabled = False
            .Buttons(cEsci).Enabled = False
            lstODL.Enabled = False
         End With
'      Case "C"
'         iRes = MsgBox("Vuoi chiudere il corrente reclamo", _
'               vbYesNo + vbDefaultButton2 + vbQuestion)
'         If iRes = vbYes Then
'            sqlc = "Update ReclamoVenditore set Stato = 'N' " & _
'                     "Where ODL = '" & Me.Key
'            db.Execute (sqlc)
'            tbCommand.Buttons(cChiudi).Enabled = False
'         End If
   End Select
End Sub


