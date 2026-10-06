VERSION 5.00
Begin VB.Form frmImposta 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Impostazione"
   ClientHeight    =   4440
   ClientLeft      =   6480
   ClientTop       =   4410
   ClientWidth     =   6330
   Icon            =   "frmImposta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   6330
   Begin VB.TextBox txtLinea 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "Salva"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtId 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmImposta.frx":0CCA
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox txtOrdine 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtSettimana 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "Esci"
      Default         =   -1  'True
      Height          =   615
      Left            =   5520
      Picture         =   "frmImposta.frx":0CD0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Calandra"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Ordine n."
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Settimana"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Image imgLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmImposta.frx":0E1A
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Programmazione ordine"
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
      Height          =   1260
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   5505
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Qta"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   3600
      Width           =   375
   End
End
Attribute VB_Name = "frmImposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String
Dim fSalva As Boolean
Private mValore As String, sVal As String, sDato As String


Public Property Let Key(ByVal vValue As String)
   mKey = Trim$(vValue)
 
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub SalvaID()
Dim sqlc As String, rsId As Recordset, scod As Variant
    
    sqlc = "SELECT ID, Settimana, Ordinamento, Linea " & _
            "From Formule " & _
            "WHERE " & sDato
    Set rsId = db.OpenRecordset(sqlc, dbOpenDynaset)
   Do While Not rsId.EOF
        rsId.Edit
        rsId("Settimana") = txtSettimana.Text
        rsId("Ordinamento") = txtOrdine.Text
        rsId("Linea") = txtLinea.Text
        rsId.Update
        rsId.MoveNext
    Loop


End Sub

Private Sub cmdEsci_Click()
 Unload Me
End Sub

Private Sub lblRisposta_Click()

End Sub

Private Sub cmdSalva_Click()
Controllo
If fSalva Then
    SpostaId
End If

End Sub

Private Sub Form_Load()
'txtId.Text = Me.Key
txtSettimana.Text = Date
txtLinea.Text = 1
txtOrdine.Text = 1
End Sub
Private Sub Controllo()
fSalva = False
If Not IsNull(txtSettimana.Text) And txtSettimana.Text <> "" Then
    If Not IsNull(txtOrdine.Text) And txtOrdine.Text <> "" Then
        fSalva = True
    Else
        MsgBox "Prima di salvare un ordine bisogna inserire l'ordine di programmazione", vbInformation
        fSalva = False
        txtOrdine.SetFocus
    End If
Else
    MsgBox "Prima di salvare un ordine bisogna inserire la settimana di programmazione", vbInformation
    fSalva = False
    txtSettimana.SetFocus
End If
If Not IsDate(txtSettimana.Text) Then
    MsgBox "Bisogna inserire una data valida!"
    fSalva = False
    txtSettimana.SetFocus
End If

If Not txtLinea.Text = "1" And Not txtLinea.Text = "2" Then
    MsgBox "Bisogna inserire il codice della calandra compreso tra 1 e 2"
    fSalva = False
    txtLinea.SetFocus
End If


End Sub
Private Sub SpostaId()
Dim iRes As Integer, sqlc As String
Dim iPos As Integer, sValore As String
        iRes = MsgBox("Volete mettere in preparazione l'ID" & Me.Key, _
               vbYesNo + vbDefaultButton2 + vbQuestion)
        If iRes = vbYes Then
        sValore = Me.Key
        sDato = ""
        
            sDato = "Id='" & Left$(sValore, 6) & "' "
            iPos = InStr(1, sValore, ", ")
        Do While Not iPos = 0
            If Not iPos > 6 Then
                sDato = sDato & " or Id='" & Mid$(sValore, iPos + 2, 6) & "' "
                iPos = InStr(7, sValore, ", ")
            Else
                sDato = sDato & " or Id='" & Mid$(sValore, iPos + 2, 6) & "' "
                iPos = InStr(iPos + 1, sValore, ", ")
            End If
        Loop
            
            sqlc = "Update Formule set Caricata = 'S' " & _
                     "Where Caricata = 'N' and (" & sDato & ")"
            db.Execute (sqlc)
            SalvaID
            AssegnaSe
            fSalva = False
            Unload Me
            frmTV.ChangedTV = True
        
        End If
        
End Sub

Private Sub AssegnaSe()
Dim sqlc As String, rsMp As Recordset, rsS As Recordset
Dim sMP As String, sSe As String, sOggi As String

sOggi = Format(txtSettimana.Text, "mm/dd/yyyy")
sqlc = "SELECT Formule.Mescola, Formule.ID " & _
        "From Formule " & _
        "WHERE " & sDato
Set rsMp = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsMp.EOF Then
        sMP = rsMp("Mescola")
    End If
sqlc = "SELECT Formule.Mescola, Formule.ID, Formule.Caricata, Formule.SeID, " & _
            "Formule.Settimana, Formule.Ordinamento, Formule.Linea " & _
        "From Formule " & _
        "WHERE Formule.Mescola='" & sMP & "' AND Caricata='s' AND " & _
            "Settimana=#" & sOggi & "# AND " & _
            "Ordinamento=" & txtOrdine.Text & " AND " & _
            "Linea=" & txtLinea.Text

Set rsS = db.OpenRecordset(sqlc, dbOpenDynaset)
    Do While Not rsS.EOF
        If Not IsNull(rsS("SeID")) And rsS("SeId") <> "" Then
            sSe = rsS("SeID")
        End If
    rsS.MoveNext
    Loop
    
    rsS.MoveFirst
        
    If IsNull(sSe) Or sSe = "" Then
        CreaSe
        sSe = sVal
    End If
    Do While Not rsS.EOF
        If IsNull(rsS("SeID")) Then
            rsS.Edit
            rsS("SeID") = sSe
            rsS.Update
        End If
        
    rsS.MoveNext
    Loop



'Chiudo i rs

rsMp.Close
Set rsMp = Nothing
rsS.Close
Set rsS = Nothing


End Sub


Private Sub CreaSe()
Dim rsNum As Recordset, sqlc As String

'   sqlc = "update UltimaRisposta " & _
               "set UltimaRisposta = UltimaRisposta + 1 " & _
               "Where dummy=0"
'   db.Execute (sqlc)
'    Dim sqlc As String, rsNum As Recordset
    sqlc = "Select max(SeId) as MSE  from SE "
    Set rsNum = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsNum.EOF Then
        If Not IsNull(rsNum("MSE")) And rsNum("MSE") <> "" Then
            sVal = rsNum("MSE")
        Else
            sVal = 0
        End If
        
    Else
        sVal = 0
    End If
    sVal = sVal + 1
    sqlc = "insert into SE " & _
            "(seID, SeS,Data, Qta)  " & _
            "values ('" & sVal & " ', 'N', Now, '0')"
            

   db.Execute (sqlc)
    
   rsNum.Close
   Set rsNum = Nothing
End Sub

Private Sub Form_Paint()
txtId.Text = Me.Key

End Sub

Private Sub Form_Terminate()
txtId.Text = Me.Key
End Sub
