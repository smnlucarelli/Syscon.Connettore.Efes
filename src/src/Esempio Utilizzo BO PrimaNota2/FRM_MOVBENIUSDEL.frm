VERSION 5.00
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.6#0"; "TMS_EDIT.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_MOVBENIUSDEL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancellazione prima nota - movimento IVA beni usati"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "FRM_MOVBENIUSDEL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Avanzamento:"
      Height          =   1695
      Left            =   60
      TabIndex        =   8
      Top             =   3420
      Width           =   8655
      Begin VB.ListBox LST_AVANZAMENTO 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   90
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   8505
      End
   End
   Begin VB.CommandButton CMD_MODIFICA 
      Caption         =   "Interroga"
      Height          =   525
      Left            =   3750
      Picture         =   "FRM_MOVBENIUSDEL.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   150
      Width           =   1305
   End
   Begin VB.CommandButton CMD_DELETE 
      Caption         =   "Cancella"
      Height          =   525
      Left            =   7440
      Picture         =   "FRM_MOVBENIUSDEL.frx":28EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2820
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   3645
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         MaxChar         =   12
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   12
         CanRequired     =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   "Numero di registrazione"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1755
      End
   End
   Begin TabDlg.SSTab TAB_GRID 
      Height          =   1965
      Left            =   90
      TabIndex        =   3
      Top             =   810
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3466
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Testata"
      TabPicture(0)   =   "FRM_MOVBENIUSDEL.frx":2A36
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GRID_TESTATA"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Movimenti contabili"
      TabPicture(1)   =   "FRM_MOVBENIUSDEL.frx":2A52
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRID_MOVCONT"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Movimenti IVA"
      TabPicture(2)   =   "FRM_MOVBENIUSDEL.frx":2A6E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GRID_MOVIVA"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "CGM1_DOCBENIUSATI"
      TabPicture(3)   =   "FRM_MOVBENIUSDEL.frx":2A8A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GRID_CGM1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "CGM2_MOVBENIUSATI"
      TabPicture(4)   =   "FRM_MOVBENIUSDEL.frx":2AA6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GRID_CGM2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin MSDataGridLib.DataGrid GRID_TESTATA 
         Height          =   1515
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   2672
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GRID_MOVCONT 
         Height          =   1515
         Left            =   -74940
         TabIndex        =   5
         Top             =   360
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   2672
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GRID_MOVIVA 
         Height          =   1515
         Left            =   -74940
         TabIndex        =   10
         Top             =   360
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   2672
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GRID_CGM1 
         Height          =   1515
         Left            =   -74940
         TabIndex        =   11
         Top             =   360
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   2672
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GRID_CGM2 
         Height          =   1515
         Left            =   -74940
         TabIndex        =   12
         Top             =   360
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   2672
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FRM_MOVBENIUSDEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Pcls_PrimaNota       As CGBO_PRIMANOTA.CLSCG_PRIMANOTA
Attribute Pcls_PrimaNota.VB_VarHelpID = -1
Public StrConnect                       As Variant
Private Connessione                     As ADODB.Connection
Public CallingForm                      As FRM_MAIN

Private Sub CMD_DELETE_Click()
    On Error GoTo Err_CMD_DELETE_Click
    
    '
    ' Cancellazione registrazione
    '
    Set Pcls_PrimaNota.ActiveInterface = CallingForm.ActiveInterface
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata
        .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
        .NumeroRegistrazione = TXT_NUMREG.Text
    End With
    
    Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
    Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione
    Pcls_PrimaNota.PGestRegPN.CancellaRegistrazione
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOk Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.CancellaRegistrazione"
        Exit Sub
    End If
    
    Pcls_PrimaNota.PGestRegPN.RegistraModifiche
    
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOk Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.RegistraModifiche"
        Exit Sub
    End If
    
    If Pcls_PrimaNota.PGestRegPN.StatoNonBloccante <> tsOk Then
        MsgBox Pcls_PrimaNota.PGestRegPN.ErroreNonBloccante
    End If
    
    MsgBox "La registrazione è stata cancellata", vbInformation
    
    Set GRID_TESTATA.DataSource = Nothing
    Set GRID_MOVCONT.DataSource = Nothing
    Set GRID_MOVIVA.DataSource = Nothing
    
Exit Sub
Err_CMD_DELETE_Click:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub CMD_MODIFICA_Click()
    On Error GoTo Err_CMD_MODIFICA_Click
    
    '
    ' Pulisco il list box che segnala le operazioni
    '
    LST_AVANZAMENTO.Clear
    LST_AVANZAMENTO.Refresh
    
    '
    ' Valorizzo le proprietà della classe che gestisce la prima nota
    '
    
    '
    ' Recupero registrazione
    '
    Set Pcls_PrimaNota.ActiveInterface = CallingForm.ActiveInterface
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata
        .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
        .NumeroRegistrazione = TXT_NUMREG.Text
    End With
    
    Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
    Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione
    Pcls_PrimaNota.PGestRegPN.ModificaRegistrazione
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOk Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.ModificaRegistrazione"
        Exit Sub
    End If
    
    Set GRID_TESTATA.DataSource = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata
    Set GRID_MOVCONT.DataSource = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont
    Set GRID_MOVIVA.DataSource = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovIva
    
    Set GRID_CGM1.DataSource = Pcls_PrimaNota.PGestRegPN.CPOutput.ClsBeniUsati.RecSetDocBeniUsati
    Set GRID_CGM2.DataSource = Pcls_PrimaNota.PGestRegPN.CPOutput.ClsBeniUsati.RecSetMovBeniUsati
    
    LST_AVANZAMENTO.AddItem String(40, "-")
    LST_AVANZAMENTO.ListIndex = LST_AVANZAMENTO.ListCount - 1
    
Exit Sub
Err_CMD_MODIFICA_Click:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    TAB_GRID.Tab = 0
    
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Form_Load
    
    Me.Left = 400
    Me.Top = 1300
    
    '
    ' Creo la connessione
    '
    Set Connessione = New ADODB.Connection
    Connessione.ConnectionString = StrConnect
    Connessione.CursorLocation = adUseClient
    Connessione.Open
    
    Set CallingForm.ActiveInterface.Connection = Connessione
    
    Set Pcls_PrimaNota = New CGBO_PRIMANOTA.CLSCG_PRIMANOTA
    
Exit Sub
Err_Form_Load:
    MsgBox Err.Number & " - " & Err.Description, , "FORM_LOAD"
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim RSOggetto       As Object
    
    On Error Resume Next
    
    If Not Pcls_PrimaNota.PGestRegPN Is Nothing Then
        '
        ' CLASSE INPUT
        '
        If Not Pcls_PrimaNota.PGestRegPN.CPInput Is Nothing Then
            If Not Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata Is Nothing Then
                Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.RecSetAssociazioneAliquote = Nothing
            End If
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata = Nothing
            
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestataPagSosp = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglioPagSosp = Nothing
            
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.RecSetImpPagAbbAcc = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.RecSetScadenze = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.RecSetPagRit = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.RecSetAltreRitenute = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.RecSetTestataIntra = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.RecSetRigheIntra = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.RecSetVenditaCespiti = Nothing
            
            If Not Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.CollectionRsetAcconti Is Nothing Then
                For Each RSOggetto In Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.CollectionRsetAcconti
                    Set RSOggetto = Nothing
                    Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.CollectionRsetAcconti.Remove 1
                Next
            End If
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.CollectionRsetAcconti = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate = Nothing
            
            Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Nothing
        End If
        
        '
        ' CLASSE OUTPUT
        '
        If Not Pcls_PrimaNota.PGestRegPN.CPOutput Is Nothing Then
            Set Pcls_PrimaNota.PGestRegPN.CPOutput.InfoDocumento.CPOutput = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPOutput.InfoDocumento.RecSetCastellettoIvaCorrente = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPOutput.InfoDocumento.RecSetCastellettoIvaOriginale = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPOutput.InfoDocumento = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPOutput.RegistrazioneFiglia = Nothing
            Set Pcls_PrimaNota.PGestRegPN.CPOutput.RegistrazionePadre = Nothing
        End If
        
        Set Pcls_PrimaNota.PGestRegPN.CPInput = Nothing
        Set Pcls_PrimaNota.PGestRegPN.CPOutput = Nothing
    End If
    
    Pcls_PrimaNota.PGestRegPN = Nothing
    Pcls_PrimaNota.CPInput = Nothing
    Pcls_PrimaNota.CPOutput = Nothing
    Pcls_PrimaNota.ActiveInterface = Nothing
    
    If Not (Connessione Is Nothing) Then
        Connessione.Close
    End If
    Set Connessione = Nothing
    
    Set CallingForm = Nothing
    
    Err.Clear
End Sub

Private Sub Pcls_PrimaNota_Avanzamento(StepAvanzamento As Integer, DescrOperazione As Variant)
    On Error GoTo Err_Pcls_GestRegPN_Avanzamento
    LST_AVANZAMENTO.AddItem DescrOperazione
    LST_AVANZAMENTO.ListIndex = LST_AVANZAMENTO.ListCount - 1
Exit Sub
Err_Pcls_GestRegPN_Avanzamento:
    MsgBox Err.Number & " - " & Err.Description, , "Pcls_PrimaNota_Avanzamento"
    Exit Sub
End Sub

'
' Funzione che sostituisce un valore nullo con un valore specificato
'
Private Function NVL(Valore As Variant, ValIfNull As Variant) As Variant
    On Error GoTo Err_NVL
    
    If IsEmpty(Valore) Or IsNull(Valore) Then
        NVL = ValIfNull
    Else
        If Trim(CStr(Valore)) = "" Then
            NVL = ValIfNull
        Else
            NVL = Valore
        End If
    End If
    
Exit Function
Err_NVL:
    MsgBox Err.Number & " - " & Err.Description, , "NVL"
    Exit Function
End Function
