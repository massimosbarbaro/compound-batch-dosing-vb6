VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTV 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   7050
   ClientLeft      =   405
   ClientTop       =   1170
   ClientWidth     =   11250
   Icon            =   "frmTV.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Stamp"
      Height          =   585
      Left            =   9240
      Picture         =   "frmTV.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   585
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Canc"
      Height          =   585
      Left            =   8520
      Picture         =   "frmTV.frx":11FC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   585
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      Height          =   585
      Left            =   10320
      Picture         =   "frmTV.frx":12FE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   585
   End
   Begin MSComctlLib.ImageList imglView 
      Left            =   10680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":1448
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":15A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":16FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":191E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   0
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
            Picture         =   "frmTV.frx":1D72
            Key             =   "D"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":21C6
            Key             =   "F"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":261A
            Key             =   "R"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":2A6E
            Key             =   "V"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":2EC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":3316
            Key             =   "C"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":3776
            Key             =   "A"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":3BD6
            Key             =   "N"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTV.frx":3D76
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvReclami 
      Height          =   3015
      Left            =   6000
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView tvReclami 
      Height          =   5655
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   9975
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmTV.frx":3F16
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblSaveODL 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblCursor 
      BackColor       =   &H000000C0&
      Height          =   3375
      Left            =   4680
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      ToolTipText     =   "Click & Drag to resize"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      Height          =   6015
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   10935
   End
   Begin VB.Label lblElenco 
      Caption         =   "Elenco di Id in produzione"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   7695
   End
   Begin VB.Label lblID 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   6360
      Width           =   7695
   End
End
Attribute VB_Name = "frmTV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lDx As Long
Private Const MinLblCursor = 75
Private lInitW As Long, lInitH As Long, lDh As Long
Private fGetFromLoad As Boolean
Private LastNode As Node
Private mChangedTV As Boolean, mParam As String
Private sNodeType As String, sWhereStato As String
' Aggiunta x Sort
Private sOrderBy As String
Private Const cDefOrderBy = " Order by 1 asc, ODL asc "
Private mValore As String, sVal As String, fSposta As Boolean
Dim sT As String, sNodeN As String

' Fine Aggiunta x Sort


Public Property Get ChangedTV() As Boolean
   ChangedTV = mChangedTV
End Property
Public Property Let ChangedTV(vValue As Boolean)
   mChangedTV = vValue
   If mChangedTV Then
      SetTV
'      Set tvReclami.Nodes.Item.Index = 91
'    tvReclami_NodeClick LastNode
     mChangedTV = False
   End If
End Property
Public Property Let Param(vValue As String)
   If vValue = "ATC" Or vValue = "MKG" Or vValue = "ATC/MKG" Or vValue = "VIEW" Then
      mParam = vValue
      With Me
         .ChangedTV = True
         .Caption = "Visualizzazione del materiale in produzione "
      End With
   Else
      Err.Raise 17, , "Errore nel passaggio del paramtero al TV"
   End If
End Property
Public Property Get Param() As String
   Param = mParam
End Property

Private Sub SetTV()
   Dim sqlc As String, rsTV As Recordset, i As Byte
   Dim nodX As Node, nodY As Node
   Dim sStato As String
   Dim sLivello2 As String, sLivello3 As String, sLivello4 As String, sLivello5 As String
   Dim rsLinea As Recordset
   Dim sLegame As String, sChiave As String, sContenuto As String, iImmagine As Integer
       
   Screen.MousePointer = vbArrowHourglass
   Set LastNode = Nothing     ' Inizializzo comunque lastnode a niente
   With tvReclami
      .Nodes.Clear
      .ImageList = ImageList1
      
'
'  Crea i Nodi Principali (Root, ATC / MKG  e Stato)
'
      i = 1
      Set nodX = .Nodes.Add(, , "R", "Produzione", i)
      
'M Costriusco il nodo Linea in automatico
If Me.Param = "ATC" Then
    sqlc = "SELECT Descrizione, Ordine " & _
            "FROM Formule INNER JOIN Descrizione " & _
                "ON Formule.Caricata = Descrizione.Caricata " & _
            "GROUP BY Descrizione, Ordine " & _
            "ORDER BY Ordine"
Else
    sqlc = "SELECT Descrizione, Ordine " & _
            "FROM Formule INNER JOIN Descrizione " & _
                "ON Formule.Caricata = Descrizione.Caricata " & _
            "GROUP BY Descrizione, Ordine " & _
            "HAVING (((Descrizione.Descrizione)='Pianificata'or (Descrizione.Descrizione)='InLavorazione')) " & _
            "ORDER BY Ordine"
End If

Set rsLinea = db.OpenRecordset(sqlc, dbOpenSnapshot)
    'Costruisco i sottonodi livello2
    Do While Not rsLinea.EOF
        sLegame = "R"
        sChiave = UCase(rsLinea(0) & "X")
        sContenuto = rsLinea(0)
        iImmagine = 5
        Set nodX = .Nodes.Add(sLegame, tvwChild, sChiave, sContenuto, iImmagine)
        rsLinea.MoveNext
    Loop
rsLinea.Close

'Inizio nodo livello Calandra______________________________________________________________________
If Me.Param = "ATC" Then
    sqlc = "SELECT [Descrizione] & 'X' AS Stato, Formule.Linea " & _
            "FROM Formule INNER JOIN Descrizione ON " & _
                "Formule.Caricata = Descrizione.Caricata " & _
            "GROUP BY [Descrizione]& 'X', Formule.Linea, " & _
                "Descrizione.Ordine " & _
            "ORDER BY Descrizione.Ordine"
Else
    
    sqlc = "SELECT [Descrizione] & 'X' AS Stato, Formule.Linea " & _
            "FROM Formule INNER JOIN Descrizione ON " & _
                "Formule.Caricata = Descrizione.Caricata " & _
            "GROUP BY Formule.Linea, Descrizione.Descrizione " & _
            "HAVING ((([Descrizione] & 'X')='PianificataX')) OR " & _
                "((([Descrizione] & 'X')='InLavorazioneX'))"

End If

    Set rsLinea = db.OpenRecordset(sqlc, dbOpenSnapshot)
'riempio il nodo famiglia
    Do While Not rsLinea.EOF
        sLegame = UCase(rsLinea("Stato"))
        sChiave = UCase(sLegame & rsLinea("Linea") & "M")
        If Not IsNull(rsLinea("Linea")) Then
            sContenuto = "Calandra " & rsLinea("Linea")
        Else
            sContenuto = "Calandra da definire"
        End If
        iImmagine = 3
        Set nodX = .Nodes.Add(sLegame, tvwChild, sChiave, sContenuto, iImmagine)
        rsLinea.MoveNext
    Loop
rsLinea.Close
'Fine nodo livello Calandra______________________________________________________________________

'Inizio nodo livello Settimana __________________________________________________________________
    
If Me.Param = "ATC" Then

    sqlc = "SELECT [Descrizione] & 'X' & [Linea] & 'M' AS Stato, Formule.Settimana " & _
            "FROM Formule INNER JOIN Descrizione ON " & _
                "Formule.Caricata = Descrizione.Caricata " & _
            "GROUP BY [Descrizione] & 'X' & [Linea] &'M', Formule.Settimana "
Else
    sqlc = "SELECT [Descrizione] & 'X' & [Linea] & 'M' AS Stato, Formule.Settimana " & _
            "FROM Formule INNER JOIN Descrizione ON " & _
                "Formule.Caricata = Descrizione.Caricata " & _
            "Where (((Descrizione.Descrizione)='Pianificata')) OR (((Descrizione.Descrizione)='InLavorazione')) " & _
            "GROUP BY [Descrizione] & 'X' & [Linea] & 'M', Formule.Settimana"
End If

    Set rsLinea = db.OpenRecordset(sqlc, dbOpenSnapshot)
'riempio il nodo famiglia
    Do While Not rsLinea.EOF
        sLegame = UCase(rsLinea("Stato"))
        sChiave = UCase(sLegame & rsLinea("Settimana") & "S")
        If Not IsNull(rsLinea("Settimana")) Then
            sContenuto = rsLinea("Settimana")
        Else
            sContenuto = "Settimana da definire"
        End If
        iImmagine = 4
        Set nodX = .Nodes.Add(sLegame, tvwChild, sChiave, sContenuto, iImmagine)
        rsLinea.MoveNext
    Loop
rsLinea.Close
'Fine nodo livello Settimana __________________________________________________________________

'Inizio nodo livello Ordinamento_______________________________________________________________
    
If Me.Param = "ATC" Then

    sqlc = "SELECT [Descrizione] & 'X' & [Linea] & 'M' & [Settimana] & 'S' AS Stato, " & _
                "Formule.Ordinamento " & _
            "FROM Formule INNER JOIN Descrizione ON " & _
                "Formule.Caricata = Descrizione.Caricata " & _
            "GROUP BY [Descrizione] & 'X' & [Linea] & 'M' & [Settimana] & 'S' , " & _
                "Formule.Ordinamento "
Else
    sqlc = "SELECT [Descrizione] & 'X' & [Linea] & 'M' & [Settimana] & 'S'  AS Stato, " & _
                "Formule.Ordinamento " & _
            "FROM Formule INNER JOIN Descrizione ON " & _
                "Formule.Caricata = Descrizione.Caricata " & _
            "Where (((Descrizione.Descrizione) = 'Pianificata' or (Descrizione.Descrizione) = 'InLavorazione')) " & _
            "GROUP BY [Descrizione] & 'X' & [Linea] & 'M' & [Settimana] & 'S' , " & _
                "Formule.Ordinamento "
End If

    Set rsLinea = db.OpenRecordset(sqlc, dbOpenSnapshot)
'riempio il nodo famiglia
    Do While Not rsLinea.EOF
        sLegame = UCase(rsLinea("Stato"))
        sChiave = UCase(sLegame & rsLinea("Ordinamento") & "O")
        If Not IsNull(rsLinea("Ordinamento")) Then
            sContenuto = rsLinea("Ordinamento")
        Else
            sContenuto = "Ordinamento da definire"

        End If
        iImmagine = 8
        Set nodX = .Nodes.Add(sLegame, tvwChild, sChiave, sContenuto, iImmagine)
        rsLinea.MoveNext
    Loop
rsLinea.Close
'Fine nodo livello Ordinamento_______________________________________________________________

'Inizio nodo livello Mescole_________________________________________________________________

If Me.Param = "ATC" Then
    sqlc = "SELECT [Descrizione].[Descrizione] & 'X' & [Linea] & 'M' & [Settimana] & 'S' & [Ordinamento] & 'O' AS Stato, " & _
                "UCase([Mescola] & ' - ' & [descrizmesc].[Descrizione]) AS Mesc " & _
            "FROM (Formule LEFT JOIN Descrizione ON Formule.Caricata = " & _
                "Descrizione.Caricata) LEFT JOIN DescrizMesc ON Formule.Mescola = DescrizMesc.MP " & _
            "GROUP BY [Descrizione].[Descrizione] & 'X' & [Linea] & 'M' & [Settimana] & 'S' & [Ordinamento] & 'O'," & _
            " UCase([Mescola] & ' - ' & [descrizmesc].[Descrizione])"
Else
    sqlc = "SELECT [Descrizione].[Descrizione] & 'X' & [Linea] & 'M' & [Settimana] " & _
                "& 'S' & [Ordinamento] & 'O' AS Stato, UCase([Mescola] " & _
                "& ' - ' & [descrizmesc].[Descrizione]) AS Mesc " & _
            "FROM (Formule LEFT JOIN Descrizione ON Formule.Caricata " & _
                "= Descrizione.Caricata) LEFT JOIN DescrizMesc ON " & _
                "Formule.Mescola = DescrizMesc.MP " & _
            "WHERE ((([Descrizione].[Descrizione] & 'X')='PianificataX' or ([Descrizione].[Descrizione] & 'X')='InLavorazioneX')) " & _
            "GROUP BY [Descrizione].[Descrizione] & 'X' & [Linea] " & _
                "& 'M' & [Settimana] & 'S' & [Ordinamento] & 'O', " & _
                "UCase([Mescola] & ' - ' & [descrizmesc].[Descrizione])"
End If

    Set rsLinea = db.OpenRecordset(sqlc, dbOpenSnapshot)
'riempio il nodo famiglia
    Do While Not rsLinea.EOF
        sLegame = UCase(rsLinea("Stato"))
        sChiave = UCase(sLegame & rsLinea("Mesc") & "E")
        If Not IsNull(rsLinea("Mesc")) Then
            sContenuto = rsLinea("Mesc")
        Else
            sContenuto = "Mescola da definire"
        End If
        iImmagine = 6
        Set nodX = .Nodes.Add(sLegame, tvwChild, sChiave, sContenuto, iImmagine)
        rsLinea.MoveNext
    Loop
rsLinea.Close
'Fine nodo livello Mescole_________________________________________________________________

     Set rsLinea = Nothing


'      .BorderStyle = 1
      
      .BorderStyle = vbFixedSingle
      .Checkboxes = False
      .FullRowSelect = True
      .HideSelection = False
      .HotTracking = True
      .Indentation = 100
      .LabelEdit = tvwManual
      .LineStyle = tvwTreeLines
      .Scroll = True
'      .SingleSel = True   Non so se mi piace !!
      'Se allo Style non metti niente il programma funziona con + e -
      .Style = tvwTreelinesPlusMinusPictureText
      
      With .Nodes("R")
         .Expanded = True
         .BackColor = vbYellow
         .Bold = True
      End With
'      With .Nodes("ATC")
'         .Expanded = True
'         .Bold = True
'         .BackColor = vbBlue
'      End With
'      With .Nodes("MKG")
'         .Expanded = True
'         .Bold = True
'         .BackColor = vbGreen
'      End With
'      With .Nodes("NA")
'         .Bold = True
'         .ForeColor = vbRed
'         .Selected = True
'      End With
   End With
'
'  Imposta la Header del List View
'
    lblID.Caption = "" 'pulisco lalista di Id da preparare
   SetLVHeader
   Screen.MousePointer = vbNormal

'   tvReclami.Nodes("NA").Selected = True
End Sub
Private Sub SetLVHeader()
   Dim clX As ColumnHeader
   With lvReclami
      .ColumnHeaders.Clear
      .View = lvwReport
      .Icons = imglView
      .SmallIcons = imglView
      Select Case Me.Param
         Case "ATC"
            Set clX = .ColumnHeaders.Add(, , "ID", 1000, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Data", 1000, lvwColumnCenter)
            Set clX = .ColumnHeaders.Add(, , "ODL", 1000, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "SeId", 700, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "CodArt", 1565, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Qta", 1000, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Mescola", 700, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "IdMescola", 1000, lvwColumnRight)
         Case Else
            Set clX = .ColumnHeaders.Add(, , "ID", 1000, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "Data", 1000, lvwColumnCenter)
            Set clX = .ColumnHeaders.Add(, , "ODL", 1000, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "SeId", 700, lvwColumnLeft)
            Set clX = .ColumnHeaders.Add(, , "CodArt", 1565, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Qta", 1000, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "Mescola", 700, lvwColumnRight)
            Set clX = .ColumnHeaders.Add(, , "IdMescola", 1000, lvwColumnRight)
        End Select
      .Arrange = lvwAutoLeft
   End With
End Sub

Private Sub cmdDel_Click()
sT = ""
lblID.Caption = ""
End Sub

Private Sub cmdEsci_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Me.PrintForm
End Sub

Private Sub Form_Activate()
   Dim nodX As Node
'Per Giudo Così Funziona
   With tvReclami
         On Error Resume Next
   
      If .SelectedItem <> "" Then
'      On Error Resume Next
         Set nodX = .Nodes(.SelectedItem.Index)
      Else
         Set nodX = .Nodes(1)    '  Nuovi
      End If
   End With
   tvReclami_NodeClick nodX
End Sub

Private Sub Form_Click()
'   Dim i As Integer
'   Dim strNodes As String
'   For i = 1 To tvReclami.Nodes.Count
'   strNodes = strNodes & tvReclami.Nodes(i).Index & " " & _
'               "Key: " & tvReclami.Nodes(i).Key & " " & _
'               "Text: " & tvReclami.Nodes(i).Text & vbLf
'   Next i
'   MsgBox strNodes
End Sub


Private Sub Form_Load()
   Dim sqlc As String
   
   Screen.MousePointer = vbArrowHourglass
   
   fGetFromLoad = True
   fSposta = False
   
   lblSaveODL.Visible = False
   lDx = tvReclami.Left - Shape1.Left
   With Me
      lDh = Shape1.Top
      .Height = lDh + Shape1.Height + 2 * lDh + cmdEsci.Height + _
            (.Height - .ScaleHeight)
      .Width = Shape1.Left + Shape1.Width + 2 * (.Width - .ScaleWidth)
      lInitW = .Width
      lInitH = .Height
   End With
   fGetFromLoad = False
'
'   Me.ChangedTV = True
'
'   Me.Show
'
'  Simulo un Click sul nodo "Nuovi"
'
'   Set nodX = tvReclami.Nodes("N")
'   tvReclami_NodeClick nodX
   Screen.MousePointer = vbNormal
' Aggiunta x Sort
' "order by  rv.odl asc, Soluzione asc, DataReclamo desc "
   sOrderBy = cDefOrderBy
' Fine Aggiunta x Sort
'nascondo gli elenchi degli id
    Elenco False


End Sub

Private Sub Form_Resize()
   If Not fGetFromLoad Then
      With Me
         If .Width < lInitW Then .Width = lInitW
         If .Height < lInitH Then .Height = lInitH
         Shape1.Width = .ScaleWidth - 2 * Shape1.Left
         Shape1.Height = .ScaleHeight - (3 * lDh + cmdEsci.Height)
         lvReclami.Width = Shape1.Width - lDx - lvReclami.Left
      End With
      With tvReclami
         .Height = Shape1.Height - (2 * (tvReclami.Top - Shape1.Top))
      End With
      With lvReclami
         .Left = tvReclami.Left + tvReclami.Width + MinLblCursor
         .Top = tvReclami.Top
         .Height = tvReclami.Height
         .Width = Shape1.Width + Shape1.Left - lDx - .Left
      End With
      With cmdEsci
         .Top = Me.ScaleHeight - (.Height + lDh)
         .Left = Shape1.Left + Shape1.Width - .Width
      End With
      With cmdDel
         .Top = Me.ScaleHeight - (.Height + lDh)
         .Left = Shape1.Left + lblID.Width + .Width + 50
      End With
      With cmdPrint
         .Top = Me.ScaleHeight - (.Height + lDh)
         .Left = Shape1.Left + Shape1.Width - .Width - cmdDel.Width - 200
      End With
      
      With lblID
         .Top = Me.ScaleHeight - (.Height + lDh - 90)
         .Left = Shape1.Left '+ Shape1.Width - Shape1.Width / 2 - cmdDel.Width * 2
      End With
      With lblElenco
         .Top = Me.ScaleHeight - (.Height + lDh + 500)
         .Left = Shape1.Left '+ Shape1.Width - .Width - Shape1.Width / 2 - cmdDel.Width * 2
      End With
'      With lblExplain
'         .Top = cmdEsci.Top + (cmdEsci.Height - .Height) \ 2
'         .Left = lvReclami.Left + (lvReclami.Width - .Width) \ 2
'      End With
   End If
   With lblCursor
'      .Height = tvReclami.Height * 1.01
     .Height = tvReclami.Height * 0.99
     .Width = (tvReclami.Width / 2) + (lvReclami.Width / 2) + MinLblCursor
     .Top = tvReclami.Top * 1.01
     .Left = tvReclami.Left + (tvReclami.Width / 2)
     .BackColor = Shape1.BackColor
   End With
End Sub

Private Sub lblCursor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      If X < MinLblCursor Then
         X = MinLblCursor
      End If
      If X > lblCursor.Width Then
         X = lblCursor.Width
      End If
      lvReclami.Left = lblCursor.Left + X
      lvReclami.Width = Shape1.Width - lDx - lvReclami.Left
      tvReclami.Width = (lblCursor.Left + X) - tvReclami.Left - MinLblCursor
   End If
End Sub

Private Sub lvReclami_BeforeLabelEdit(Cancel As Integer)
'   MsgBox "Non è possibile modificare i valori.", vbOKOnly + vbExclamation, ""
'   Beep
   Cancel = True
End Sub


Private Sub lvReclami_DblClick()
   Dim sEntrata As String, sPos As String, iRes As Integer
   Dim sqlc As String, rsT As Recordset, sSE As String
    If lblSaveODL.Caption <> "" Then
        Select Case sNodeType
            Case "Mescola"
                sPos = InStr(1, mValore, "X")
                sEntrata = Left$(mValore, sPos - 1)
                Select Case sEntrata
                    Case "NUOVA"
                        SpostaId
                        If fSposta Then
                            Load frmImposta
                            '                        frmImposta.Key = lblSaveODL.Caption
                            frmImposta.Key = sT
                            frmImposta.Show vbModal
                            sT = ""
                        End If
                    Case "DAPIANIFICARE"
                        SpostaId
                        If fSposta Then
                            Load frmID
                            frmID.Key = lblSaveODL.Caption & ", " & mValore
                            frmID.Show
                        End If
                    Case "PIANIFICATA"
                        If Not Me.Param = "ATC" Then
 '                           MsgBox "Sposta in lavorazione"
                            iRes = MsgBox("Vuoi iniziare questa produzione?", _
                            vbYesNo + vbDefaultButton2 + vbQuestion)
                            If iRes = vbYes Then
                                sSE = ""
                                sqlc = "SELECT ID, SeID " & _
                                        "From Formule " & _
                                        "WHERE ID='" & lblSaveODL.Caption & "'"
                                Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
                                    If Not rsT.EOF Then
                                        sSE = rsT("SeId")
                                    End If
                                rsT.Close
                                    
                                sqlc = "Update Formule set Caricata = 'L' " & _
                                        "Where SeId = " & sSE & " and " & _
                                        "Caricata = 'P'"
                                db.Execute (sqlc)
                            End If
                            'Inserire il tvchange
                            frmTV.ChangedTV = True
                              Pulisci

                        Else
                            SpostaId
'                            MsgBox "Sposta in lavorazione"
                            iRes = MsgBox("Vuoi iniziare questa produzione?", _
                            vbYesNo + vbDefaultButton2 + vbQuestion)
                            If iRes = vbYes Then
                                sSE = ""
                                sqlc = "SELECT ID, SeID " & _
                                        "From Formule " & _
                                        "WHERE ID='" & lblSaveODL.Caption & "'"
                                Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
                                    If Not rsT.EOF Then
                                        sSE = rsT("SeId")
                                    End If
                                rsT.Close
                                    
                                sqlc = "Update Formule set Caricata = 'L' " & _
                                        "Where SeId = " & sSE & " and " & _
                                        "Caricata = 'P'"
                                db.Execute (sqlc)
                            End If
                            'Inserire il tvchange
                            frmTV.ChangedTV = True
                              Pulisci
                        End If
                    'questo sotto lo tolgo perchè l'ho messo in lavorazone
                    
                    '                        If fSposta Then
                    '                            Load frmBilance
                    '                            frmBilance.Key = lblSaveODL.Caption & ", " & mValore
                    '                            frmBilance.Show
                    '                        End If
                    '    MsgBox "Non puoi modificare tu lo stato della materia " & _
                    '        "con la cuasale pianificata"
                     '
                    '    frmTV.ChangedTV = True

                    Case "INLAVORAZIONE"
                    
                    '                       SpostaId    'tolgo lo sposta perchè oramai la produzione è cominciata
'                        If fSposta Then
                            Load frmBilance
                            frmBilance.Key = lblSaveODL.Caption & ", " & mValore
                            frmBilance.Show
 '                       End If
                    Case "NONTRASFERITA"
                        Load frmID
                        frmID.Key = lblSaveODL.Caption & ", " & mValore
                        frmID.Show
                    Case "TRASFERITA"
                        Load frmID
                        frmID.Key = lblSaveODL.Caption & ", " & mValore
                        frmID.Show
                        '       MsgBox "FineTrasferita"
                End Select
            Case "Princ"
                sPos = InStr(1, mValore, "X")
                sEntrata = Left$(mValore, sPos - 1)

                If sEntrata = "TRASFERITA" Then
                    MsgBox "vuoi stampare"
                
                                        Load frmID
                        frmID.Key = lblSaveODL.Caption & ", " & mValore
                        frmID.Show

                
                '''
                '
                'DEvo inserire il menù di stampa della mescola
                '
                '
                '
                '
                '
                '
                '
                Else
                    MsgBox "Selezionare prima il nome di una mescola.", _
                    vbOKOnly + vbExclamation
                
                End If
                
            
            
            Case Else
                MsgBox "Selezionare prima il nome di una mescola.", _
                vbOKOnly + vbExclamation
        End Select
    Else
       MsgBox "Nessun ODL Selezionato.", _
             vbOKOnly + vbExclamation
    End If
End Sub

Private Sub lvReclami_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Stop
' Aggiunta x Sort
    Select Case ColumnHeader.Index
        Case 1, 2  ' ordinamenti possibili: colonne 1 e 2 (ODL, Data)
            sOrderBy = " Order by " & Trim$(Str$(ColumnHeader.Index)) & " "
'   Se si vuole lasciare l'ordinamento anche per le altre colonne
        Case 3 To 100  ' ordinamenti possibili: colonne 1 e 2 (ODL, Data)
            sOrderBy = " Order by " & Trim$(Str$(ColumnHeader.Index)) & " "
        Case Else   ' se seleziono un'altra colonna --> reset ordine iniziale (1)
            sOrderBy = cDefOrderBy
    End Select
    'Modifica per errore icona nodo
    
'    tvReclami_NodeClick LastNode
    tvReclami_NodeClick tvReclami.SelectedItem
'   Fine modifica per errore icona
' Fine Aggiunta x Sort
    
End Sub

Private Sub lvReclami_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim sI As String
    lblSaveODL.Caption = Trim$(Item.Text)
    
    If Len(Item.Text) = 5 Then
        sI = Item.Text & " "
    Else
        sI = Item.Text
    End If
    
    If sNodeType = "Mescola" And sNodeN = "NUOVA" Then



        sT = SetId(sT, sI)

'        If Not Trim$(Right(sT, 6)) = sI Or Not (Len(sT) < 7) Then
'            If Not sT = "" Then
'                sT = sT & ", " & sI
'            Else
'                sT = sI
'            End If
'        Else
'            sT = sI
'        End If
        
        
        lblID.Visible = True
        lblElenco.Visible = True
        cmdDel.Visible = True
        lblID.Caption = sT
        lblID.Visible = True
'        If Len(sT) < 7 Then
'        Stop
'        End If
        
    End If
End Sub



Private Sub tvReclami_Collapse(ByVal Node As MSComctlLib.Node)
   If Node.Index = 1 Then
      Node.Expanded = True   ' Expand the node again.
      tvReclami_NodeClick Node
   End If
End Sub

Private Sub tvReclami_Expand(ByVal Node As MSComctlLib.Node)
'   Debug.Print "Expanded " & Node.Text
End Sub

Private Sub tvReclami_NodeClick(ByVal Node As MSComctlLib.Node)
   Dim sqlc As String, rsLV As Recordset, iImg As Integer
   Dim itmX As ListItem, i As Integer
   Dim nodeX As Node
   Dim sTabella As String
   lblSaveODL.Caption = ""    '  Elimino l'ultimo odl selezionato
'
'   Pulisco comunque la parte ListView
'
    lvReclami.ListItems.Clear
    sNodeN = ""
   If Node.Index = 1 Then  ' Si tratta del nodo Root
'      Set nodeX = tvReclami.Nodes("NA")
'      nodeX.Selected = True
'      tvReclami_NodeClick nodeX
   Else
      If Node.Key = "" Then      '  Nodo Foglia
         Node.Image = "A"
         If Not (LastNode Is Nothing) Then
            If Node <> LastNode Then
               LastNode.Image = "C"
            End If
         End If
         Set LastNode = Node
'M Seleziono se fai il clic sul nodo ATC=A oppure MKG
         If Len(Node.Parent.Key) < 7 And Right$(Node.Parent.Key, 1) = "X" Then
            sNodeType = "ATC"
         Else
             mValore = Trim$(Node.Parent.Key)
          sNodeType = "Prod"
         End If
      Else
            If Node.Key <> "ATC" And Node.Key <> "MKG" Then
               Select Case Right$(Node.Key, 1)
                   Case "X" 'Principali
                       sqlc = "SELECT Formule.ID, Formule.Data, Formule.OdL, Formule.SeID, " & _
                                    "Formule.CodArt, Formule.Qta, UCase(Formule.Mescola) AS Mescola, " & _
                                    "Note.ID AS IdMescola " & _
                                "FROM ((Formule LEFT JOIN Descrizione ON Formule.Caricata " & _
                                    "= Descrizione.Caricata) LEFT JOIN DescrizMesc ON " & _
                                    "Formule.Mescola = DescrizMesc.MP) LEFT JOIN [Note] " & _
                                    "ON Formule.SeID = Note.SE " & _
                                "Where Descrizione.Descrizione='" & SetApice(Left(Trim$(Node.Key), Len(Node.Key) - 1)) & "' " & _
                                sOrderBy
                           
                       sNodeType = "Princ"
                   Case "M" 'Calandra
                       sqlc = "SELECT Formule.ID, Formule.Data, Formule.OdL, Formule.SeID, " & _
                                    "Formule.CodArt, Formule.Qta, UCase(Formule.Mescola) AS Mescola, " & _
                                    "Note.ID AS IdMescola " & _
                                "FROM ((Formule LEFT JOIN Descrizione ON Formule.Caricata " & _
                                    "= Descrizione.Caricata) LEFT JOIN DescrizMesc ON " & _
                                    "Formule.Mescola = DescrizMesc.MP) LEFT JOIN [Note] " & _
                                    "ON Formule.SeID = Note.SE " & _
                                "Where [Descrizione].[Descrizione] & 'X' & [Linea] & 'M' ='" & SetApice(Node.Key) & "' " & _
                                         sOrderBy
                           
                       sNodeType = "Linea"
                   
                   Case "S" 'Settimana
                       sqlc = "SELECT Formule.ID, Formule.Data, Formule.OdL, Formule.SeID, " & _
                                    "Formule.CodArt, Formule.Qta, UCase(Formule.Mescola) AS Mescola, " & _
                                    "Note.ID AS IdMescola " & _
                                "FROM ((Formule LEFT JOIN Descrizione ON Formule.Caricata " & _
                                    "= Descrizione.Caricata) LEFT JOIN DescrizMesc ON " & _
                                    "Formule.Mescola = DescrizMesc.MP) LEFT JOIN [Note] " & _
                                    "ON Formule.SeID = Note.SE " & _
                               "Where [Descrizione].[Descrizione] & 'X' & [Linea] & 'M' & [Settimana] & 'S' ='" & SetApice(Node.Key) & "' " & _
                               sOrderBy
                               
                       sNodeType = "Settimana"
                   
                   Case "O" 'Ordinamento
                       sqlc = "SELECT Formule.ID, Formule.Data, Formule.OdL, Formule.SeID, " & _
                                    "Formule.CodArt, Formule.Qta, UCase(Formule.Mescola) AS Mescola, " & _
                                    "Note.ID AS IdMescola " & _
                                "FROM ((Formule LEFT JOIN Descrizione ON Formule.Caricata " & _
                                    "= Descrizione.Caricata) LEFT JOIN DescrizMesc ON " & _
                                    "Formule.Mescola = DescrizMesc.MP) LEFT JOIN [Note] " & _
                                    "ON Formule.SeID = Note.SE " & _
                               "Where [Descrizione].[Descrizione] & 'X' & [Linea] & 'M' & [Settimana] " & _
                                   "& 'S' & [Ordinamento] & 'O' ='" & SetApice(Node.Key) & "' " & _
                               sOrderBy
                               
                               
                       sNodeType = "Ordinamento"
                   
                   Case "E" 'Mescola
                       sqlc = "SELECT Formule.ID, Formule.Data, Formule.OdL, Formule.SeID, " & _
                                    "Formule.CodArt, Formule.Qta, UCase(Formule.Mescola) AS Mescola, " & _
                                    "Note.ID AS IdMescola " & _
                                "FROM ((Formule LEFT JOIN Descrizione ON Formule.Caricata " & _
                                    "= Descrizione.Caricata) LEFT JOIN DescrizMesc ON " & _
                                    "Formule.Mescola = DescrizMesc.MP) LEFT JOIN [Note] " & _
                                    "ON Formule.SeID = Note.SE " & _
                               "WHERE [Descrizione].[Descrizione] & 'X' & [Linea] " & _
                                   "& 'M' & [Settimana] & 'S' & [Ordinamento] & 'O' & " & _
                                   "[Mescola] & ' - ' & [DescrizMesc].[Descrizione] & 'E'='" & SetApice(Node.Key) & "' " & _
                               sOrderBy
                       sNodeType = "Mescola"
                   Case Else
                       sqlc = ""
               End Select
               mValore = Trim$(Node.Key)
            Else
               sqlc = ""
            End If
            sNodeN = Left$(Node.Key, 5)
      End If


      With lvReclami
         If sqlc <> "" Then
            Set rsLV = db.OpenRecordset(sqlc, dbOpenDynaset)
            Do While Not rsLV.EOF
               Select Case rsLV("Mescola")
                  Case "C"
                     iImg = 1
                  Case "N"
                     iImg = 2
                  Case "V"
                     iImg = 3
                  Case Else
                    iImg = 3
               End Select
' Carico una icona differente per le segnalazioni
   
               'If rsLV("Soluzione") = 2 Then
               '   iImg = 4
               'End If
'**************************************************************
               Set itmX = .ListItems.Add(, , rsLV(0), , iImg)
               With itmX
'M Seleziono chi entra per far vedere una cosa o l'altra
            Select Case Me.Param
                Case "ATC"
                    sTabella = rsLV.Fields.Count - 1
                    
                    Select Case sNodeType
                        Case "Prod"
                        sTabella = rsLV.Fields.Count - 1
                        Case "Linea"
                        sTabella = rsLV.Fields.Count - 1
                        Case "Famiglia"
                        sTabella = rsLV.Fields.Count - 1
                    End Select
                  Case Else
                    If sNodeType <> "Prod" Then
                        sTabella = rsLV.Fields.Count - 1
                    Else
                        sTabella = rsLV.Fields.Count - 5
                    End If
                
            End Select
                For i = 1 To sTabella
                     If Not IsNull(rsLV(i)) Then
                        itmX.SubItems(i) = CStr(rsLV(i))
                     Else
                        itmX.SubItems(i) = " "
                     End If
                  Next i
               .Ghosted = False
               
               End With
               rsLV.MoveNext
            Loop
            rsLV.Close
            Set rsLV = Nothing
         End If
      End With
   End If
   
' Modifica per errore icona nel nodo
'Set LastNode = Node
'   Fine modifica
'nascondo gli elenchi degli id
    Elenco False


End Sub
Private Sub SpostaId()
'Si vuole fare in modo che al doppio clic su un Id di una mescola si sposti da
'un punto di patenza ad un punto di arrivo
Dim sqlc As String, rsId As Recordset
Dim iRes As Integer
        fSposta = False
'         MsgBox "Alla prossima domanda rispondi sempre no!!!!! " & vbCrLf & "Ciao Massimo"
       
        iRes = MsgBox("Volete modificare la programmazione dell'ID" & lblSaveODL.Caption, _
               vbYesNo + vbDefaultButton2 + vbQuestion)
        If iRes = vbYes Then
        'apro la finestra di programmazione
            Load frmSposta
            frmSposta.Key = lblSaveODL.Caption
            frmSposta.Show vbModal
            Pulisci
'            Me.Status = cExists
'            frmTV.ChangedTV = True
        Else
            fSposta = True
        End If

    
    
End Sub

Private Sub AssegnaSe()
'Dim sqlc As String, rsMp As Recordset, rsS As Recordset
'Dim sMP As String, sSe As String'

'sqlc = "SELECT Formule.Mescola, Formule.ID " & _
'        "From Formule " & _
'        "WHERE ID = '" & lblSaveODL.Caption & "'"
'Set rsMp = db.OpenRecordset(sqlc, dbOpenSnapshot)
'    If Not rsMp.EOF Then
'        sMP = rsMp("Mescola")
'    End If
'sqlc = "SELECT Formule.Mescola, Formule.ID, Formule.Caricata, Formule.SeID " & _
'        "From Formule " & _
'        "Where Formule.Mescola = '" & sMP & "' And Formule.Caricata = 's'"'

'Set rsS = db.OpenRecordset(sqlc, dbOpenDynaset)
'    Do While Not rsS.EOF
'        If Not IsNull(rsS("SeID")) And rsS("SeId") <> "" Then
'            sSe = rsS("SeID")
'        End If
'    rsS.MoveNext
'    Loop
'
'    rsS.MoveFirst
'
'    If IsNull(sSe) Or sSe = "" Then
'        CreaSe
'        sSe = sVal
'    End If
'    Do While Not rsS.EOF
'        If IsNull(rsS("SeID")) Then
'            rsS.Edit
'            rsS("SeID") = sSe
'            rsS.Update
'        End If
'
'    rsS.MoveNext
'    Loop'
''
''
'
''Chiudo i rs'
'
'rsMp.Close
'Set rsMp = Nothing
'rsS.Close
'Set rsS = Nothing'
'
'
End Sub


Private Sub CreaSe()


'Dim rsNum As Recordset, sqlc As String''

'    sqlc = "Select max(SeId) as MSE  from SE "
'    Set rsNum = db.OpenRecordset(sqlc, dbOpenSnapshot)
'    If Not rsNum.EOF Then
'        If Not IsNull(rsNum("MSE")) And rsNum("MSE") <> "" Then
'            sVal = rsNum("MSE")
'        Else
'            sVal = 0
'        End If
'
'    Else
'        sVal = 0
'    End If
'    sVal = sVal + 1
'    sqlc = "insert into SE " & _
'            "(seID, SeS,Data, Qta)  " & _
'            "values ('" & sVal & " ', 'N', Now, '0')"
''
'
 '  db.Execute (sqlc)
 '
'   rsNum.Close
'   Set rsNum = Nothing
End Sub
Private Sub Elenco(Flag As Boolean)

'nascondo gli elenchi degli id
    lblID.Visible = Flag
    lblElenco.Visible = Flag
    cmdDel.Visible = Flag


End Sub
Private Sub Pulisci()
   Dim nodX As Node
'Per Giudo Così Funziona
   With tvReclami
         On Error Resume Next
   
      If .SelectedItem <> "" Then
'      On Error Resume Next
         Set nodX = .Nodes(.SelectedItem.Index)
      Else
         Set nodX = .Nodes(1)    '  Nuovi
      End If
   End With
   tvReclami_NodeClick nodX

End Sub
