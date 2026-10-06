VERSION 5.00
Begin VB.Form frmSposta 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Sposta"
   ClientHeight    =   4215
   ClientLeft      =   4620
   ClientTop       =   3900
   ClientWidth     =   6090
   Icon            =   "frmSposta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6090
   Begin VB.CommandButton cmdSalva 
      Cancel          =   -1  'True
      Caption         =   "Sposta"
      Default         =   -1  'True
      Height          =   615
      Left            =   360
      Picture         =   "frmSposta.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Elimina"
      Height          =   615
      Left            =   1320
      Picture         =   "frmSposta.frx":11FC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtCalandra 
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtOrdinamento 
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtSettimana 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtMescola 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   615
      Left            =   5160
      Picture         =   "frmSposta.frx":172E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   615
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
      Height          =   345
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmSposta.frx":1878
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Label Label4 
      Caption         =   "Id n."
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Mescola n."
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Settimana n."
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Ordine n."
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Spostamento ordine"
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
      TabIndex        =   3
      Top             =   240
      Width           =   4545
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmSposta.frx":187E
      Top             =   0
      Width           =   450
   End
   Begin VB.Label Label5 
      Caption         =   "Calandra n."
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   4695
   End
End
Attribute VB_Name = "frmSposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String
Dim fSalva As Boolean
Private mValore As String, sVal As String, sDato As String
Private sSE As String, fAM As Boolean


Public Property Let Key(ByVal vValue As String)
   mKey = Trim$(vValue)
 
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub SalvaID()
Dim sqlc As String, rsId As Recordset, scod As Variant
    
    
sqlc = "SELECT Mescola, Data, OdL, ID, CodArt, " & _
            "Qta, Caricata, SeID, Settimana, Ordinamento, " & _
            "Linea " & _
        "From Formule " & _
        "Where ID = '" & Me.Key & "'"

Set rsId = db.OpenRecordset(sqlc, dbOpenDynaset)
    Do While Not rsId.EOF
        rsId.Edit
            If Not IsNull(txtOrdinamento.Text) And txtOrdinamento.Text <> "" Then
                rsId("Ordinamento") = txtOrdinamento.Text
            Else
                rsId("Ordinamento") = ""
            End If
            If Not IsNull(txtSettimana.Text) And txtSettimana.Text <> "" Then
                rsId("Settimana") = txtSettimana.Text
            Else
                rsId("Settimana") = ""
            End If
            
            If Not IsNull(txtCalandra.Text) And txtCalandra.Text <> "" Then
                rsId("Linea") = txtCalandra.Text
            Else
                rsId("Linea") = ""
            End If
        rsId.Update
        rsId.MoveNext
    Loop

End Sub

Private Sub cmdDel_Click()
 EliminaId

End Sub

Private Sub cmdEsci_Click()
 Unload Me
End Sub

Private Sub cmdSalva_Click()
Controllo
If fSalva Then
    SpostaId
End If

End Sub

Private Sub Form_Load()
'txtId.Text = Me.Key
End Sub
Private Sub Controllo()
fSalva = False
If Not IsNull(txtSettimana.Text) And txtSettimana.Text <> "" Then
    If Not IsNull(txtOrdinamento.Text) And txtOrdinamento.Text <> "" Then
        fSalva = True
    Else
        MsgBox "Prima di salvare un ordine bisogna inserire l'ordine di programmazione", vbInformation
        fSalva = False
        txtOrdinamento.SetFocus
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

If Not txtCalandra.Text = "1" And Not txtCalandra.Text = "2" Then
    MsgBox "Bisogna inserire il codice della calandra compreso tra 1 e 2"
    fSalva = False
    txtCalandra.SetFocus
Else
    fSalva = True
End If


End Sub
Private Sub SpostaId()
Dim iRes As Integer, sqlc As String
Dim iPos As Integer, sValore As String
        iRes = MsgBox("Volete realmente spostare l'ID" & Me.Key, _
               vbYesNo + vbDefaultButton2 + vbQuestion)
        If iRes = vbYes Then
            sValore = Me.Key
            sDato = ""
            sDato = "Id='" & Trim$(sValore) & "' "
            SalvaID
            AssegnaSe
            fSalva = False
            Unload Me
            frmTV.ChangedTV = True
 
        End If
 
End Sub

Private Sub AssegnaSe()
Dim sqlc As String, rsMp As Recordset, rsS As Recordset
Dim sMP As String, sSE As String, sOggi As String

sqlc = "SELECT Mescola, ID, Caricata, SeID, " & _
            "Settimana, Ordinamento, Linea " & _
        "From Formule " & _
        "WHERE " & sDato & " AND " & _
            "Linea=" & txtCalandra.Text '

Set rsS = db.OpenRecordset(sqlc, dbOpenDynaset)
        CreaSe
        sSE = sVal

    Do While Not rsS.EOF
        rsS.Edit
        rsS("SeID") = sSE
        rsS.Update
    rsS.MoveNext
    Loop



rsS.Close
Set rsS = Nothing


End Sub


Private Sub CreaSe()
Dim rsNum As Recordset, sqlc As String

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
'txtId.Text = Me.Key
TrovaId
End Sub

Private Sub TrovaId()
Dim sqlc As String, rsT As Recordset

sqlc = "SELECT Mescola, Data, OdL, ID, CodArt, " & _
            "Qta, Caricata, SeID, Settimana, Ordinamento, " & _
            "Linea " & _
        "From Formule " & _
        "Where ID = '" & Me.Key & "'"

Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
If Not rsT.EOF Then
    If Not IsNull(rsT("Id")) And rsT("Id") <> "" Then
        txtId.Text = rsT("Id")
        txtId.Locked = True
    Else
        txtId.Text = ""
        txtId.Locked = True
    End If
    If Not IsNull(rsT("Mescola")) And rsT("Mescola") <> "" Then
        txtMescola.Text = Format(rsT("Mescola"), ">")
        txtMescola.Locked = True
    Else
        txtMescola.Text = ""
        txtMescola.Locked = True
    End If
    If Not IsNull(rsT("Settimana")) And rsT("Settimana") <> "" Then
        txtSettimana.Text = rsT("Settimana")
    Else
        txtSettimana.Text = ""
    End If
    
    If Not IsNull(rsT("Ordinamento")) And rsT("Ordinamento") <> "" Then
        txtOrdinamento.Text = rsT("Ordinamento")
    Else
        txtOrdinamento.Text = ""
    End If

    If Not IsNull(rsT("Linea")) And rsT("Linea") <> "" Then
        txtCalandra.Text = rsT("Linea")
    Else
        txtCalandra.Text = ""
    End If
    If Not IsNull(rsT("SeId")) And rsT("SeId") <> "" Then
        sSE = rsT("SeId")
    Else
        sSE = ""
    End If
   
End If

End Sub
Private Sub EliminaId()
Dim iRes As Integer
Dim sqlc As String, rsE As Recordset

        iRes = MsgBox("Volete realmente eliminare l'ID" & Me.Key, _
               vbYesNo + vbDefaultButton2 + vbQuestion)
        If iRes = vbYes Then
                sqlc = "DELETE Mescola, Data, OdL, ID, CodArt, Qta, " & _
                            "Caricata, SeID, Settimana, Ordinamento, Linea " & _
                        "From Formule " & _
                        "WHERE ID='" & Me.Key & "'"
                
                db.Execute (sqlc)
            fSalva = False
            Unload Me
            frmTV.ChangedTV = True
            ApriMateriali

        End If

End Sub
Private Sub ApriMateriali()
Dim sqlc As String, rsAM As Recordset
Dim sM As String, iRes As Integer, sWhere As String
fAM = False
If Not IsNull(sSE) And Not sSE = "" Then
        sWhere = " SeID= " & sSE
Else
        sWhere = " ID = '" & Me.Key & "'"
End If

sqlc = "SELECT Mescola, Data, OdL, ID, CodArt, " & _
            "Qta, Caricata, SeID, Settimana, Ordinamento, " & _
            "Linea " & _
        "From Formule " & _
        "Where " & sWhere
Set rsAM = db.OpenRecordset(sqlc, dbOpenDynaset)

If Not rsAM.EOF Then
    iRes = MsgBox("Avete modificato la quantità dell'ordine SEId" & sSE & vbCrLf _
            & " della mescola" & rsAM("Mescola") & "volete aprire la finestra delle mescole?", _
           vbYesNo + vbDefaultButton2 + vbQuestion)
    If iRes = vbYes Then
        sM = rsAM("ID")
        Load frmBilance
        frmBilance.Key = sM
        frmBilance.Show
    End If
End If



End Sub
