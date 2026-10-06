VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGrafici 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Grafici"
   ClientHeight    =   6255
   ClientLeft      =   4950
   ClientTop       =   2880
   ClientWidth     =   5325
   Icon            =   "frmGrafici.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   5325
   Begin VB.ComboBox cmbMesi 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdRicerche 
      Caption         =   "Ricerche rete"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccess 
      Caption         =   "Access"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "Visualizza"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   615
      Left            =   4560
      Picture         =   "frmGrafici.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   615
   End
   Begin VB.Frame fraGrafici 
      Height          =   3015
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   3375
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Esterni"
         Height          =   495
         Index           =   8
         Left            =   360
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Rete"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "&Calandra"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "&Mescole"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   345
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Esterni"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Rete"
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo Generale"
         Height          =   495
         Index           =   6
         Left            =   360
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Soluzione"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Riepilogo per Famiglia"
         Height          =   375
         Index           =   5
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin MSMask.MaskEdBox mskDataDA 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDataA 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Da"
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "A"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Mese / Anno (4 cifre)"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Grafici"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmGrafici.frx":0E14
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmGrafici"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iIndice As Integer
'Dim ac As Access.Application


Private Sub cmdAccess_Click()
Dim ac As Access.Application

'Set ac = New Access.Application
Set ac = CreateObject("Access.Application")
' put the path to your database in hereac.
'ac.OpenCurrentDatabase ("C:\Programmi\reclami\costi.mdb")
ac.OpenCurrentDatabase ("C:\Programmi\costi\costi.mdb")
' by default, the OpenReport method of the
' DoCmd object will send the report to the printerac.
ac.DoCmd.OpenReport "rDateFamiglie", acViewPreview

' close the database
ac.CloseCurrentDatabase
End Sub

Private Sub cmdEsci_Click()
Unload Me
End Sub

Private Sub cmdGrafico_Click()
   Dim sSelFormula As String, sGrafico As String
   Dim ac As Access.Application, sReport As String
   Dim sFile As String, sWhere As String, sTabella As String

'-------------------------------------------------------------------------------
' Preparo le date limite per la stampa dei grafici

   Dim sqlc As String, rsLimiti As Recordset, i As Byte
   Dim dDataInizio, dDataFine
   Dim sTitolo As String, sTest As String, sLinea As String, sGraficoFamiglia As String, sFam As String
   Dim sReteGrafico As String, sQtaGrafico As String, sGraficiReti As String
   
   
   
   
   If mskDataA.Text = "__/____" Then
      mskDataA.Text = mskDataDA.Text
   End If
   dDataFine = Format((DateAdd("d", -1, DateAdd("m", 1, "01/" & mskDataA.Text))), "dd/mm/yyyy")
   
   sTest = Format(DateSerial(Val(Mid$(mskDataDA.Text, 4, 4)), _
                            Val(Mid$(mskDataDA.Text, 1, 2)), 1), "dd/mm/yyyy")
   dDataInizio = Format(sTest, "yyyy/mm/dd")
   'Format(DateSerial(Val(Mid$(mskDataDA.Text, 4, 4)), _
                            Val(Mid$(mskDataDA.Text, 1, 2)), 1), "dd/mm/yyyy")
   If dDataFine <= dDataInizio Then
      MsgBox "Date Errate !!"
      Exit Sub
   End If
'Fine delle date
'---------------------------------------------------------------------------------
   'Apri DB con i report
'   Set ac = New Access.Application
    Set ac = CreateObject("Access.Application")
   sFile = App.Path & "\" & "bilance.mdb"
   ac.OpenCurrentDatabase sFile
   '("C:\Programmi\reclami\costi.mdb")
'   MsgBox "Stiamo Elaborando il grafico"
     sWhere = " Between #" & dDataInizio & "# And #" & dDataFine & "#"
    Select Case iIndice
     Case 0
        sReport = "rMescole"
        sTabella = "SELECT Formule.Mescola, DescrizMesc.Descrizione, " & _
                        "Formule.SeID, SE.Qta, Formule.Caricata, " & _
                        "Note.ID INTO [$$Mescola] " & _
                    "FROM ((Formule INNER JOIN SE ON Formule.SeID = " & _
                        "SE.SeId) LEFT JOIN DescrizMesc ON Formule.Mescola " & _
                        "= DescrizMesc.MP) LEFT JOIN [Note] ON Formule.SeID = Note.SE " & _
                    "WHERE (SE.Data) " & sWhere & _
                    "GROUP BY Formule.Mescola, DescrizMesc.Descrizione, " & _
                        "Formule.SeID, SE.Qta, Formule.Caricata, Note.ID " & _
                    "HAVING (((Formule.Caricata)='T'))"
                    
                    
     Case 1
        sReport = "rMesCal"
        sTabella = "SELECT Formule.Linea, Formule.Mescola, DescrizMesc.Descrizione, " & _
                        "Formule.SeID, SE.Qta, Formule.Caricata, Note.ID INTO [$$MescCal] " & _
                    "FROM ((Formule INNER JOIN SE ON Formule.SeID = SE.SeId) " & _
                        "LEFT JOIN DescrizMesc ON Formule.Mescola = DescrizMesc.MP) " & _
                        "LEFT JOIN [Note] ON Formule.SeID = Note.SE " & _
                    "WHERE (SE.Data) " & sWhere & _
                    "GROUP BY Formule.Linea, Formule.Mescola, DescrizMesc.Descrizione, " & _
                        "Formule.SeID, SE.Qta, Formule.Caricata, Note.ID " & _
                    "HAVING (((Formule.Caricata)='T'))"

                    
     Case 3
'        sReport = "rGEsterno"
      Case 4
'       sReport = "rTotaleRiepilogo"
     Case 5
'       sReport = "rTotaleRiepilogoFamiglie"
     Case 6
'       sReport = "rTotaleRiepilogoRisposta"
    Case 7
'        sReport = "rTotaleRiepilogoRete"
    Case 8
'        sReport = "rTotaleRiepilogoEsterni"
   End Select
   
   
      If iIndice < 9 Then
         'Mando in stampa il report selezionato
        ac.DoCmd.RunSQL sTabella
        ac.DoCmd.OpenReport sReport, acViewPreview
        ac.DoCmd.Maximize
        ac.Visible = True
      Else
            If Err.Number <> 0 Then
               MsgBox "Errore nel produrre il report " & vbCrLf & _
                   Err.Number & " - " & Err.Description
               Err.Clear
            End If
      End If

   
End Sub
Private Sub EseguiQuery()
'Dim q As QueryDef, db As Database, sqlc As String
'
' On Error Resume Next
' Set db = CurrentDb
'    If Err.Number <> 0 Then
'
'        CurrentDb.Close
''        db.Close
''        ac.CloseCurrentDatabase
'
'        Set db = CurrentDb
'    End If
'
'        sqlc = "Drop table  Linea"
'        On Error Resume Next
'        db.Execute (sqlc)
'        On Error Resume Next
'
'    '    On Error GoTo 0
'        sqlc = "drop table finalefamiglia"
'        On Error Resume Next
'        db.Execute (sqlc)
'        On Error Resume Next
'
''    On Error GoTo 0
''    sqlc = "Drop table $$Famiglia"
''    On Error Resume Next
''    db.Execute (sqlc)
''    On Error GoTo 0
'
'
'
'
'    Set q = db.QueryDefs("qQtaLinea1")
'    q.Execute
'    Set q = db.QueryDefs("qQtaLinea2")
'    q.Execute
'    Set q = db.QueryDefs("qQtaLinea3")
'    q.Execute
'    Set q = db.QueryDefs("qQtaLinea4")
'    q.Execute
'    Set q = db.QueryDefs("qFinaleFamiglia")
'    q.Execute
'
'
'
End Sub

Private Sub cmdRicerche_Click()
'         Load frmRicerca
'         frmRicerca.Show

End Sub

Private Sub Form_Load()
iIndice = -99
   Dim i As Byte
   For i = 1 To 12
      cmbMesi.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
   Next i
   cmbMesi.Text = Format(Month(Now), "mmmm")
   mskDataDA.Text = Format(Now, "mm/yyyy")
   Me.Show
'   mskDataDA.SetFocus
    optGrafico(0) = True
End Sub

Private Sub mskDataA_Validate(Cancel As Boolean)
   If mskDataA.Text <> "__/____" Then
      If Not IsDate("01/" & mskDataA.Text) Then
         MsgBox "Data A Non valida"
         Cancel = True
      End If
   Else
      mskDataA.Text = mskDataDA.Text
   End If
End Sub

Private Sub mskDataDA_Validate(Cancel As Boolean)
'  La DATADA NON può essere nulla !!
'   If mskDataDA.Text <> "__/____" Then
   If Not IsDate("01/" & mskDataDA.Text) Then
      MsgBox "Data DA Non valida"
      Cancel = True
   End If
'   End If
End Sub


Private Sub optGrafico_Click(Index As Integer)
iIndice = Index

End Sub

