VERSION 5.00
Begin VB.Form frmTurno 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Turno"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   Icon            =   "frmTurno.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmTurno.frx":0CCA
   ScaleHeight     =   5730
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Cancel          =   -1  'True
      Caption         =   "Stampa"
      Default         =   -1  'True
      Height          =   615
      Left            =   5280
      Picture         =   "frmTurno.frx":100C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   655
   End
   Begin VB.TextBox txtRisposta 
      Height          =   1215
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1080
      Width           =   5400
   End
   Begin VB.TextBox txtT 
      Height          =   285
      Left            =   5520
      TabIndex        =   8
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   615
      Left            =   6120
      Picture         =   "frmTurno.frx":1676
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   615
   End
   Begin VB.ComboBox cmbOperaio 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox txtMescole 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   3690
      Width           =   1455
   End
   Begin VB.CommandButton cmdAggiungi 
      Caption         =   "Aggiungi"
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Operaio"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   5895
   End
   Begin VB.Label Label3 
      Caption         =   "Turno"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Data"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmTurno.frx":17C0
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Turno"
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
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Mescole n."
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   5895
   End
End
Attribute VB_Name = "frmTurno"
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
   Turno
 
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub SalvaID()
Dim sqlc As String, rsId As Recordset, scod As Variant
Dim iPos As Integer
'   mKey = Trim$(vValue)
   Dim rsRisposta As Recordset, i As Integer
   sqlc = "SELECT Turno.Turno, Turno.IdOperaio, Turno.Mescole, " & _
                "Turno.Data, Turno.SeId " & _
            "From Turno " & _
            "Where Turno.SeId =" & Me.Key
    Set rsId = db.OpenRecordset(sqlc, dbOpenDynaset)
    
    rsId.AddNew
    rsId("SeId") = mKey
    

    scod = cmbOperaio.Text
    iPos = InStr(scod, ",")
    If iPos <> 0 Then
        scod = Left$(scod, iPos - 1)
    End If

    rsId("IdOperaio") = scod
    rsId("Turno") = txtT.Text
    rsId("Mescole") = txtMescole.Text
    rsId("Data") = txtData.Text
    rsId.Update


End Sub

Private Sub cmdAggiungi_Click()
Dim iRes As Integer, sqlc As String

Controllo
If fSalva Then
    iRes = MsgBox("Siete sicuri di voler inserire un nuovo turno?", _
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

Private Sub cmdPrint_Click()
    Me.PrintForm

End Sub

Private Sub Form_Load()
    ListCodiceOperaio
'    Turno
End Sub
Private Sub ListCodiceOperaio()
   Dim rsMp As Recordset, sqlc As String, sMsg As String
'   sqlc = "SELECT MP, Descrizione " & _
            "From DescrizMesc " & _
            "WHERE MP Like 'm*'"
    sqlc = "SELECT IdOperaio, Cognome, Nome " & _
            "From Operaio"
   Set rsMp = db.OpenRecordset(sqlc, dbOpenSnapshot)
   cmbOperaio.Clear
   Do While Not rsMp.EOF
      cmbOperaio.AddItem rsMp(0) & ", " & rsMp(2) & " " & rsMp(1)
      rsMp.MoveNext
   Loop
   rsMp.Close
   Set rsMp = Nothing
   cmbOperaio.ListIndex = -1
End Sub


Private Sub Turno()
Dim sqlc As String, rsT As Recordset, sTurno As String
Dim fT As Boolean
txtRisposta.Enabled = False
fT = False
sqlc = "SELECT Operaio.Cognome, Operaio.Nome, " & _
            "Turno.Turno, Turno.Mescole, Turno.Data, Turno.SeId " & _
        "FROM Operaio INNER JOIN Turno ON " & _
            "Operaio.IdOperaio = Turno.IdOperaio " & _
        "WHERE Turno.SeId=" & mKey
        
    Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
        Do While Not rsT.EOF
            If Not fT Then
                sTurno = rsT("Cognome") & ", " & rsT("Nome") & " Turno " & rsT("Turno") _
                        & ", " & rsT("Data") & " Mescole n. " & rsT("Mescole")
                fT = True
            Else
                sTurno = sTurno & vbCrLf & rsT("Cognome") & ", " & rsT("Nome") & " Turno " & rsT("Turno") _
                        & ", " & rsT("Data") & " Mescole n. " & rsT("Mescole")
            End If
            rsT.MoveNext
        Loop
        
        txtRisposta.Text = sTurno
End Sub
Private Sub Controllo()
fSalva = False
If Not IsDate(txtData.Text) Then
    MsgBox "Bisogna inserire una data valida!"
    fSalva = False
    txtData.SetFocus
Else
    If Not IsNumeric(txtT.Text) Then
        MsgBox "Bisogna inserire un numero corrispondente al turno!"
        fSalva = False
        txtT.SetFocus
    Else
        If Not IsNumeric(txtMescole.Text) Then
            MsgBox "Bisogna inserire il numero di mescole girate nel turno!"
            fSalva = False
            txtMescole.SetFocus
        Else
            fSalva = True
        End If
    End If
End If


End Sub

