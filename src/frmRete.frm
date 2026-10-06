VERSION 5.00
Begin VB.Form frmAggiunte 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Componenti"
   ClientHeight    =   3510
   ClientLeft      =   5805
   ClientTop       =   1875
   ClientWidth     =   6510
   Icon            =   "frmRete.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   6510
   Begin VB.CommandButton cmdAggiungi 
      Caption         =   "Aggiungi"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtMpQ 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox cmbMp 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1440
      Width           =   2625
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "Esci"
      Default         =   -1  'True
      Height          =   615
      Left            =   5400
      Picture         =   "frmRete.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Materia da aggiungere"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Qta"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aggiunta componenti"
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
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   4725
   End
   Begin VB.Image imgLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmRete.frx":0E14
      Top             =   0
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   4695
   End
End
Attribute VB_Name = "frmAggiunte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String
Dim fSalva As Boolean

Public Property Let Key(ByVal vValue As String)
Dim sqlc As String, rsRisposta As Recordset, i As Integer
   
   mKey = Trim$(vValue)
   
   sqlc = "SELECT MP.SeID, MP.MP, MP.Lotto, MP.Qta " & _
            "From MP " & _
            "WHERE MP.SeId=" & Me.Key
   
'   sqlc = "SELECT ATC.NumeroRisposta, ATC.DataRisposta " & _
          "FROM ATC where NumeroRisposta = " & Me.Key
   Set rsRisposta = db.OpenRecordset(sqlc, dbOpenSnapshot)
   
   
   
'   For i = 0 To rsRisposta.Fields.Count - 1
'      If Not IsNull(rsRisposta(i)) Then
'         lblDettagli(i).Caption = rsRisposta(i)
'      Else
'         lblDettagli(i) = ""
'      End If
'   Next i
   
 
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub SalvaID()
Dim sqlc As String, rsId As Recordset, scod As Variant

'   mKey = Trim$(vValue)
   Dim rsRisposta As Recordset, i As Integer
   sqlc = "SELECT SeID, MP, Lotto, Qta, QtaC, QtaT, QtaP " & _
            "From MP " & _
            "WHERE SeId=" & Me.Key
   
    Set rsId = db.OpenRecordset(sqlc, dbOpenDynaset)
    
    rsId.AddNew
    rsId("SeId") = mKey
    scod = SalvaCodiceMp(cmbMp.Text)
    rsId("Mp") = scod
    rsId("Lotto") = "0"
    rsId("Qta") = "0"
    rsId("QtaC") = "0"
    rsId("QtaP") = "0"
    rsId("QtaT") = txtMpQ.Text
    rsId.Update


End Sub

Private Sub cmdAggiungi_Click()
Dim iRes As Integer, sqlc As String

Controllo
If fSalva Then
    iRes = MsgBox("Siete sicuri di voler modificare la mescola?", _
        vbYesNo + vbDefaultButton2 + vbQuestion)
        If iRes = vbYes Then
            SalvaID
            Unload Me
        End If
End If

End Sub

Private Sub cmdEsci_Click()
 Unload Me
End Sub

Private Sub lblRisposta_Click()

End Sub

Private Sub Form_Load()
    ListCodiceMP
End Sub
Private Sub ListCodiceMP()
   Dim rsMp As Recordset, sqlc As String, sMsg As String
   sqlc = "SELECT MP, Descrizione " & _
            "From DescrizMesc " & _
            "WHERE MP Like 'm*'"
   Set rsMp = db.OpenRecordset(sqlc, dbOpenSnapshot)
   cmbMp.Clear
   Do While Not rsMp.EOF
      cmbMp.AddItem rsMp(0) & ", " & rsMp(1)
      rsMp.MoveNext
   Loop
   rsMp.Close
   Set rsMp = Nothing
   cmbMp.ListIndex = -1
End Sub
Private Function SalvaCodiceMp(s As String)
   Dim rsMp As Recordset, sqlc As String, sDescrizione As String
   'Dim iPos As Integer
   'iPos = InStr(s, "'")
   'If iPos <> 0 Then
   '   s = Left$(s, iPos) & "'" & Right$(s, Len(s) - iPos)
   '
   'End If
   s = Left$(s, 5)
   sqlc = "SELECT MP " & _
            "From DescrizMesc " & _
            "WHERE MP='" & Trim$(s) & "'"
'   sqlc = "SELECT Codmp " & _
           "FROM mp where Descrizionemp='" & Trim$(s) & "'"
   Set rsMp = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsMp.EOF Then
      SalvaCodiceMp = rsMp("Mp")
   Else
      SalvaCodiceMp = ""
   End If
   rsMp.Close
   Set rsMp = Nothing
End Function
Private Sub Controllo()
fSalva = False
If Not IsNull(txtMpQ.Text) And txtMpQ.Text <> "" Then
    fSalva = True
Else
    MsgBox "Prima di aggiungere un componenete bisogna inserire la quantità", vbInformation
    fSalva = False
End If

    
End Sub

