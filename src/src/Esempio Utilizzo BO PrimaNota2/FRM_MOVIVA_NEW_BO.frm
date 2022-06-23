VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_COMBOBOX.ocx"
Object = "{D8EB97B9-26FF-11D2-A1C0-0060082875F9}#6.17#0"; "TMS_EDITDEFCONTO.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.9#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.8#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.7#0"; "TMS_EDITDATE.ocx"
Begin VB.Form FRM_MOVIVA_NEW_BO 
   Caption         =   "Form1"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMD_INSERT 
      Caption         =   "Esegui"
      Height          =   525
      Left            =   8820
      Picture         =   "FRM_MOVIVA_NEW_BO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6090
      Width           =   1305
   End
   Begin VB.Frame Frame8 
      Caption         =   "Avanzamento:"
      Height          =   1755
      Left            =   60
      TabIndex        =   48
      Top             =   6720
      Width           =   10095
      Begin VB.ListBox LST_AVANZAMENTO 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   90
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   9945
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Dati relativi al dettaglio:"
      Height          =   1515
      Left            =   60
      TabIndex        =   30
      Top             =   4530
      Width           =   10095
      Begin PRJFW_EDITM.TXT_EDITM TXT_ALIQUOTA1 
         Height          =   300
         Left            =   2160
         TabIndex        =   31
         Top             =   570
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         IsLookup        =   -1  'True
         DisplayFormat   =   "Maiuscolo"
         MaxChar         =   4
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         NumRighe        =   0
         MaxWidth        =   4
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1 
         Height          =   300
         Left            =   2160
         TabIndex        =   32
         Top             =   240
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTOIVA 
         Height          =   300
         Left            =   2160
         TabIndex        =   33
         Top             =   1080
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNOIVA 
         Height          =   315
         Left            =   7620
         TabIndex        =   35
         Top             =   1080
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPONIBILE1 
         Height          =   300
         Left            =   4860
         TabIndex        =   34
         Top             =   210
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO1 
         Height          =   315
         Left            =   7620
         TabIndex        =   47
         Top             =   210
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin VB.Label Label34 
         Caption         =   "Imposta"
         Height          =   225
         Left            =   4080
         TabIndex        =   46
         Top             =   630
         Width           =   735
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPOSTA1 
         Height          =   300
         Left            =   4860
         TabIndex        =   45
         Top             =   570
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   90
         X2              =   10000
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label36 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   675
         Left            =   120
         TabIndex        =   44
         Top             =   210
         Width           =   315
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOIVA 
         Height          =   300
         Left            =   4860
         TabIndex        =   43
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label41 
         Caption         =   "Segno"
         Height          =   225
         Left            =   6990
         TabIndex        =   42
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label42 
         Caption         =   "Cod. aliquota"
         Height          =   225
         Left            =   510
         TabIndex        =   41
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label Label43 
         Caption         =   "Importo"
         Height          =   225
         Left            =   4080
         TabIndex        =   40
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label Label44 
         Caption         =   "Conto di contropartita IVA"
         Height          =   225
         Left            =   120
         TabIndex        =   39
         Top             =   1140
         Width           =   2025
      End
      Begin VB.Label Label45 
         Caption         =   "Imponibile"
         Height          =   225
         Left            =   4080
         TabIndex        =   38
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label46 
         Caption         =   "Conto di contropartita"
         Height          =   225
         Left            =   510
         TabIndex        =   37
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label47 
         Caption         =   "Segno"
         Height          =   225
         Left            =   6990
         TabIndex        =   36
         Top             =   1140
         Width           =   675
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Dati relativi alla testata:"
      Height          =   1305
      Left            =   60
      TabIndex        =   11
      Top             =   3180
      Width           =   10095
      Begin PRJFW_EDITM.TXT_EDITM TXT_CAUSALE 
         Height          =   300
         Left            =   4560
         TabIndex        =   12
         Top             =   300
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         IsLookup        =   -1  'True
         DisplayFormat   =   "Maiuscolo"
         MaxChar         =   4
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         NumRighe        =   0
         MaxWidth        =   4
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATADOC 
         Height          =   300
         Left            =   7680
         TabIndex        =   13
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREG 
         Height          =   300
         Left            =   7680
         TabIndex        =   14
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMDOCORIG 
         Height          =   300
         Left            =   4560
         TabIndex        =   15
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNOCLIFOR 
         Height          =   315
         Left            =   7680
         TabIndex        =   29
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin VB.Label Label49 
         Caption         =   "Segno"
         Height          =   225
         Left            =   6930
         TabIndex        =   28
         Top             =   990
         Width           =   675
      End
      Begin PRJFW_EDIT.TxtEdit TXT_CLIFOR 
         Height          =   300
         Left            =   1590
         TabIndex        =   27
         Top             =   960
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   529
         MaxChar         =   5
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   5
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_TOTALE 
         Height          =   300
         Left            =   4560
         TabIndex        =   26
         Top             =   960
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMDOC 
         Height          =   300
         Left            =   1590
         TabIndex        =   25
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   1590
         TabIndex        =   24
         Top             =   300
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   5
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   5
         CanRequired     =   0   'False
      End
      Begin VB.Label Label50 
         Caption         =   "Data doc."
         Height          =   225
         Left            =   6930
         TabIndex        =   23
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label51 
         Caption         =   "Num. doc. orig."
         Height          =   225
         Left            =   3390
         TabIndex        =   22
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label52 
         Caption         =   "Cliente/forn."
         Height          =   225
         Left            =   90
         TabIndex        =   21
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label53 
         Caption         =   "Codice ditta"
         Height          =   225
         Left            =   90
         TabIndex        =   20
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label54 
         Caption         =   "Data reg."
         Height          =   225
         Left            =   6930
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label55 
         Caption         =   "Causale"
         Height          =   225
         Left            =   3390
         TabIndex        =   18
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label56 
         Caption         =   "Num. doc. iniziale"
         Height          =   225
         Left            =   90
         TabIndex        =   17
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label57 
         Caption         =   "Totale fattura"
         Height          =   225
         Left            =   3390
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      Height          =   3045
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   10125
      Begin VB.ListBox TXT_TIMER_REG 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   2565
         Left            =   4320
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   1785
      End
      Begin VB.CommandButton CMD_CLEAR 
         Caption         =   "Cancella"
         Height          =   255
         Left            =   9300
         TabIndex        =   5
         Top             =   120
         Width           =   795
      End
      Begin VB.ListBox TXT_TIMER_REGMOD 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   2565
         Left            =   7980
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   1785
      End
      Begin VB.ListBox TXT_TIMER_RIGA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   2565
         Left            =   6150
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   1785
      End
      Begin VB.CheckBox CHK_DACONS_IVA 
         Caption         =   "Da consolidare"
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Tempo complessivo:"
         Height          =   225
         Left            =   180
         TabIndex        =   55
         Top             =   1740
         Width           =   1845
      End
      Begin VB.Label LBL_TEMPO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   150
         TabIndex        =   54
         Top             =   2010
         Width           =   2745
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMEROMOVIMENTI 
         Height          =   300
         Left            =   2610
         TabIndex        =   6
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin VB.Label Label26 
         Caption         =   "Numero di registrazioni da inserire"
         Height          =   225
         Left            =   180
         TabIndex        =   10
         Top             =   420
         Width           =   2445
      End
      Begin VB.Label Label27 
         Caption         =   "reg. mod."
         Height          =   225
         Left            =   7980
         TabIndex        =   9
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label28 
         Caption         =   "ins. registrazione"
         Height          =   225
         Left            =   4320
         TabIndex        =   8
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label29 
         Caption         =   "ins. riga"
         Height          =   225
         Left            =   6150
         TabIndex        =   7
         Top             =   120
         Width           =   885
      End
   End
   Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
      Height          =   300
      Left            =   90
      TabIndex        =   53
      Top             =   6360
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   12
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      MaxWidth        =   12
      CanRequired     =   0   'False
   End
   Begin VB.Label Label58 
      Caption         =   "Ultimo numero registrazione assegnato"
      Height          =   225
      Left            =   90
      TabIndex        =   52
      Top             =   6120
      Width           =   2745
   End
   Begin PRJFW_EDIT.TxtEdit TXT_PROG 
      Height          =   300
      Left            =   3000
      TabIndex        =   51
      Top             =   6090
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   6
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      MaxWidth        =   6
      CanRequired     =   0   'False
   End
End
Attribute VB_Name = "FRM_MOVIVA_NEW_BO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TipoRegistroEnum
    NessunRegistro = 0
    RegistroAcquisti = 1
    RegistroVendite = 2
    RegistroCorrispettivi = 3
End Enum

Private WithEvents Pcls_PrimaNota       As CGBO_PRIMANOTA.CLSCG_PRIMANOTA
Attribute Pcls_PrimaNota.VB_VarHelpID = -1
Public StrConnect                       As Variant
Private Connessione                     As ADODB.Connection
Public CallingForm                      As FRM_MAIN
Private Pcls_Decode                     As COBO_LOOKUPDECODE.CLSCO_DECODE
Private ClsMovGen                       As CGUO_MOVCONTABILI.CLSCG_MOVCONTABILI

Private Timer_ini                       As Single
Private Timer_fin                       As Single

Private Timer_diff_ins_reg              As Single
Private Timer_diff_ins_riga             As Single
Private Timer_diff_registramodifiche    As Single
Private Timer_diff_modificareg          As Single

Private Pobj_PNota                      As CLSCG_PNOTACHECK
Private PClsScadenze                    As PFUO_SCADENZE.CLSPF_SCADENZE

Private Sub CMD_CLEAR_Click()
    On Error Resume Next
    
    TXT_TIMER_REG.Clear
    TXT_TIMER_RIGA.Clear
    TXT_TIMER_REGMOD.Clear
    
    Err.Clear
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    TXT_DITTA.Text = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    '
    ' Setto le proprietà dei conti
    '
    SettaProprietaConti
    
    '
    ' inserimento fatt. acq. intra
    '
    TXT_NUMEROMOVIMENTI.Text = 1
    
    TXT_NUMDOC.Text = "123"
    TXT_CLIFOR.Text = 1 ' 3 = fornitore intra
    TXT_CAUSALE.Text = "VIMP"
    TXT_CAUSALE_Validate False
    TXT_NUMDOCORIG.Text = "orig"
    TXT_TOTALE.Text = 1200
    TXT_DATAREG.Text = "13/10/2012"
    TXT_DATADOC.Text = "12/10/2012"
    
    TXT_CONTO1.Text = "1200040100"
    TXT_IMPONIBILE1.Text = 1000
    TXT_ALIQUOTA1.Text = "20"
    TXT_ALIQUOTA1_Validate False
    
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
    
    Set Pcls_Decode = New COBO_LOOKUPDECODE.CLSCO_DECODE
    Set Pcls_Decode.ActiveInterface = CallingForm.ActiveInterface
    
    '
    ' Fattura acquisto INTRA
    '
    CBO_SEGNOCLIFOR.AddItemData "Dare", 1
    CBO_SEGNOCLIFOR.AddItemData "Avere", 2
    
    CBO_SEGNO1.AddItemData "Dare", 1
    CBO_SEGNO1.AddItemData "Avere", 2
    
    CBO_SEGNOIVA.AddItemData "Dare", 1
    CBO_SEGNOIVA.AddItemData "Avere", 2
    
Exit Sub
Err_Form_Load:
    MsgBox Err.Number & " - " & Err.Description, , "FORM_LOAD"
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim RSOggetto       As Object
    
    On Error Resume Next
    
    Set Pcls_Decode = Nothing
    Set Pobj_PNota = Nothing
    Set ClsMovGen = Nothing
    
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
    
    '
    ' 03.02.2010 Fabio:
    '
    Exit Sub
    
    LST_AVANZAMENTO.AddItem DescrOperazione
    LST_AVANZAMENTO.ListIndex = LST_AVANZAMENTO.ListCount - 1
Exit Sub
Err_Pcls_GestRegPN_Avanzamento:
    MsgBox Err.Number & " - " & Err.Description, , "Pcls_PrimaNota_Avanzamento"
    Exit Sub
End Sub

Private Sub SettaProprietaConti()
    On Error GoTo Err_SettaProprietaConti
    
    Set TXT_CONTO1.Connessione = Connessione
    TXT_CONTO1.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTO1.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    Set TXT_CONTOIVA.Connessione = Connessione
    TXT_CONTOIVA.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTOIVA.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
Exit Sub
Err_SettaProprietaConti:
    MsgBox Err.Number & " - " & Err.Description, , "SettaProprietaConti"
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

Private Sub DistruggiBOPrimaNota()
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
    
    Set Pcls_PrimaNota = Nothing
    
    Err.Clear
End Sub

Private Sub TXT_ALIQUOTA1_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    Dim Pcls_Lookup As CGBO_LOOKUPDECODE.CLSCG_LOOKUP
    
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Lookup = New CGBO_LOOKUPDECODE.CLSCG_LOOKUP
    
    Pcls_Lookup.CodiciIva
    
    str_SQL = Pcls_Lookup.StringaSQL
    Arr_Fields = Pcls_Lookup.ColonneLookup
    Str_Caption = Pcls_Lookup.Caption
    Str_Connect = StrConnect
    
    Set Pcls_Lookup = Nothing
    
    Err.Clear
End Sub

Private Sub TXT_ALIQUOTA1_Validate(Cancel As Boolean)
    On Error Resume Next
    AggiornaImposte
    Err.Clear
End Sub

Private Sub TXT_CAUSALE_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    Dim Pcls_Lookup As CGBO_LOOKUPDECODE.CLSCG_LOOKUP
        
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Lookup = New CGBO_LOOKUPDECODE.CLSCG_LOOKUP
    
    Pcls_Lookup.CausaliContabili
    
    str_SQL = Pcls_Lookup.StringaSQL
    Arr_Fields = Pcls_Lookup.ColonneLookup
    Str_Caption = Pcls_Lookup.Caption
    Str_Connect = StrConnect
    
    Set Pcls_Lookup = Nothing
    
    Err.Clear
End Sub

Private Sub TXT_CAUSALE_Validate(Cancel As Boolean)
    Dim TipoRegistro        As TipoRegistroEnum
    
    On Error Resume Next
    
    TipoRegistro = GetTipoRegistro(TXT_CAUSALE.Text)
    Select Case TipoRegistro
        Case TipoRegistroEnum.RegistroVendite
            CBO_SEGNOCLIFOR.Text = 1 'dare
            CBO_SEGNO1.Text = 2 'avere
            CBO_SEGNOIVA.Text = 2 'avere
            
            TXT_CONTOIVA.Text = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoIvaVendite
        Case TipoRegistroEnum.RegistroAcquisti
            CBO_SEGNOCLIFOR.Text = 2 'avere
            CBO_SEGNO1.Text = 1 'dare
            CBO_SEGNOIVA.Text = 1 'dare
            
            TXT_CONTOIVA.Text = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoIvaAcquisti
    End Select
    Err.Clear
End Sub

Private Sub TXT_CONTO1_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CONTO1_StartLookup
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CONTO1_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CONTO1_StartLookup"
    Exit Sub
End Sub

Private Sub TXT_CONTOIVA_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CONTOIVA_StartLookup
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CONTOIVA_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CONTOIVA_StartLookup"
    Exit Sub
End Sub

'
' Riempie la stringa passata con "0" a sinistra (serve per formattare i conti)
'
Private Function Fill0(Stringa As Variant, Lunghezza As Integer) As String
    Dim StrTemp     As String
    
    On Error GoTo Err_Fill0
    
    If IsNull(Stringa) Or IsEmpty(Stringa) Then
        StrTemp = String(Lunghezza, "0")
    Else
        StrTemp = Right(String(Lunghezza, "0") & Trim(CStr(Stringa)), Lunghezza)
    End If
    Fill0 = StrTemp
Exit Function
Err_Fill0:
    MsgBox Err.Number & " - " & Err.Description, , "Fill0"
    Fill0 = ""
    Err.Clear
End Function

'
' Riempie la stringa passata con spazi a destra (serve per i campi char del DB)
'
Private Function BFill(Pvar_Stringa As Variant, Pint_Len As Integer) As Variant
    Dim Pvar_Temp   As Variant
    
    On Error GoTo Err_BFill
    
    If IsNull(Pvar_Stringa) Or IsEmpty(Pvar_Stringa) Then
        Pvar_Temp = Null
    Else
        Pvar_Temp = Left(Trim(CStr(Pvar_Stringa)) & String(Pint_Len, " "), Pint_Len)
    End If
    
    BFill = Pvar_Temp
Exit Function
Err_BFill:
    BFill = ""
    Err.Clear
    Exit Function
End Function

Private Function GetTipoRegistro(CodiceCausale As Variant) As TipoRegistroEnum
    Dim StringaSQL      As Variant
    Dim RecSet          As ADODB.Recordset
    
    On Error GoTo Err_GetTipoRegistro
    
    StringaSQL = "SELECT CG33_INDTIPOREG" & _
                " FROM CG33_TABCAU" & _
                " WHERE CG33_CODICE = '" & BFill(CodiceCausale, 4) & "'"
    Set RecSet = Connessione.Execute(StringaSQL, , adCmdText)
    
    If RecSet.RecordCount > 0 Then
        GetTipoRegistro = RecSet.Fields("CG33_INDTIPOREG").Value
    Else
        GetTipoRegistro = NessunRegistro
    End If
    
    Set RecSet = Nothing
Exit Function
Err_GetTipoRegistro:
    GetTipoRegistro = NessunRegistro
    MsgBox Err.Number & " - " & Err.Description, , "GetTipoRegistro"
    Exit Function
End Function

Private Sub CMD_INSERT_Click()
    Dim Indice              As Variant
    Dim MastroClienti       As Variant
    Dim MastroFornitori     As Variant
    Dim TipoRegistro        As TipoRegistroEnum
    Dim RecSetScadenze      As ADODB.Recordset
    Dim StatoFinale         As StatoFinaleEnum
    Dim NumProt             As Variant
    
    Dim Timer_all_ini       As Single
    Dim Timer_all_fin       As Single
    Dim Timer_all_diff      As Single
    
    On Error GoTo Err_CMD_INSERT_Click
    
    Timer_all_diff = 0
    Timer_all_ini = Timer
    
    '
    ' Istanzio la classe per la gestione delle scadenze
    '
    If Not PClsScadenze Is Nothing Then
        Set PClsScadenze = Nothing
    End If
    Set PClsScadenze = New PFUO_SCADENZE.CLSPF_SCADENZE
    
    '
    ' Inserimento multiplo prima nota / scadenze
    '
    PClsScadenze.InserimentoMultiplo = True
    Pcls_PrimaNota.PGestRegPN.CPInput.InserimentoMultiplo = True
    
    '
    ' Pulisco il list box che segnala le operazioni
    '
    LST_AVANZAMENTO.Clear
    LST_AVANZAMENTO.Refresh
    
    '
    ' Pulisco il text che contiene l'ultimo numero di registrazione assegnato
    '
    TXT_NUMREG.Text = ""
    
    '
    ' Protocollo di partenza
    '
    NumProt = NVL(TXT_NUMDOC.Text, 0)
    
    '
    ' Determino i mastri cli/for
    '
    MastroClienti = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroClienti
    MastroFornitori = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroFornitori
    
Timer_diff_ins_reg = 0
Timer_diff_ins_riga = 0
Timer_diff_registramodifiche = 0
    
    For Indice = 1 To NVL(TXT_NUMEROMOVIMENTI.Text, 0)
        '
        ' Incremento il protocollo
        '
        NumProt = NumProt + 1
        
        '
        ' Valorizzo le proprietà della classe che gestisce la prima nota
        '
        Pcls_PrimaNota.Status = tsInsert
        Set Pcls_PrimaNota.ActiveInterface = CallingForm.ActiveInterface
        Pcls_PrimaNota.CPInput.RegistraEstrattoConto = True
        Pcls_PrimaNota.CPInput.RegistraPortafoglio = True
        
        Pcls_PrimaNota.CPInput.RegistraRitenuteAcconto = False
        
        '
        ' Inserimento testata
        '
        With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata
        
            .Operatore = CallingForm.ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Codice
            
            .CodiceDitta = TXT_DITTA.Text
            .DataRegistrazione = TXT_DATAREG.Text
            .NumeroRegistrazione = ""
            .CodiceCausale = TXT_CAUSALE.Text
            .NumeroDocumento = NumProt '  TXT_NUMDOC.Text
            .NumeroPartita = NumProt ' TXT_NUMDOC.Text
            .DataRegIva = TXT_DATAREG.Text
            '
            ' In questo caso vengono trattate solo le causali 1 (fattura di vendita)
            ' e 31 (fattura di acquisto)
            '
            TipoRegistro = GetTipoRegistro(TXT_CAUSALE.Text)
            Select Case TipoRegistro
                Case RegistroAcquisti
                    .ContoCliFor = Left(MastroFornitori, 2) & Fill0(TXT_CLIFOR.Text, 8)
                Case RegistroVendite
                    .ContoCliFor = Left(MastroClienti, 2) & Fill0(TXT_CLIFOR.Text, 8)
            End Select
            
            .NumeroDocumentoOrigine = TXT_NUMDOCORIG.Text
            .DataDocumentoOrigine = TXT_DATADOC.Text
            .ImportoDocumento = TXT_TOTALE.Text
            
            If CHK_DACONS_IVA.Value = vbChecked Then
                .IndicatoreTipoMovimento = DaConsolidare
            Else
                .IndicatoreTipoMovimento = Consolidato
            End If
            
            .TipoDocumento = Pcls_PrimaNota.PGestRegPN.GetTipoDocumento(.CodiceCausale)
        End With
        
        Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
        Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione

Timer_ini = Timer
        Pcls_PrimaNota.PGestRegPN.InserisciRegistrazione
Timer_fin = Timer
Timer_diff_ins_reg = Timer_diff_ins_reg + (Timer_fin - Timer_ini)

        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRegistrazione"
            Exit Sub
        End If
        
        TXT_NUMREG.Text = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
        
        TXT_PROG.Text = Indice
        
        '
        ' Inserimento riga del cli/for (partita)
        '
        With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
            .CodiceDitta = TXT_DITTA.Text
            .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
            .Conto = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.ContoCliFor
            .Importo = TXT_TOTALE.Text
            .Segno = CBO_SEGNOCLIFOR.Text
            .IndicatoreTipoOperazione = DocumentoIVA_MovimentoPartitaDocumento
        End With
        
Timer_ini = Timer
        Pcls_PrimaNota.PGestRegPN.InserisciRiga
Timer_fin = Timer
Timer_diff_ins_riga = Timer_diff_ins_riga + (Timer_fin - Timer_ini)
        
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
            Exit Sub
        End If
        
        '
        ' Inserimento prima riga di contropartita
        '
        With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
            .CodiceDitta = TXT_DITTA.Text
            .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
            .Conto = TXT_CONTO1.Text
            .Importo = TXT_IMPONIBILE1.Text
            .Imponibile = TXT_IMPONIBILE1.Text
            .Segno = CBO_SEGNO1.Text
            .CodiceAliquota = TXT_ALIQUOTA1.Text
            .Imposta = TXT_IMPOSTA1.Text
            .ImpostaND = 0
            .IndicatoreTipoOperazione = DocumentoIVA_ContropartiteSuFattura
            .CausaleIva = Null
        End With

Timer_ini = Timer
        Pcls_PrimaNota.PGestRegPN.InserisciRiga
Timer_fin = Timer
Timer_diff_ins_riga = Timer_diff_ins_riga + (Timer_fin - Timer_ini)
        
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
            Exit Sub
        End If
        
        '
        ' Inserimento riga di contropartita IVA
        '
        If NVL(TXT_IMPORTOIVA.Text, 0) <> 0 Then
            With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
                .CodiceDitta = TXT_DITTA.Text
                .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
                .Conto = TXT_CONTOIVA.Text
                .Importo = TXT_IMPORTOIVA.Text
                .Segno = CBO_SEGNOIVA.Text
                .CodiceAliquota = Null
                .Imposta = 0
                .ImpostaND = 0
                .IndicatoreTipoOperazione = DocumentoIVA_IvaSuFatturaCorrispettivo
            End With
            
Timer_ini = Timer
            Pcls_PrimaNota.PGestRegPN.InserisciRiga
Timer_fin = Timer
Timer_diff_ins_riga = Timer_diff_ins_riga + (Timer_fin - Timer_ini)
            
            If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
                MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
                Exit Sub
            End If
        End If
        
        '
        ' Genero il recordset delle scadenze
        '
        Set RecSetScadenze = GeneraRecordsetScadenze
        
        Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.RecSetScadenze = RecSetScadenze
        
        ' -----------------------------------------------------
        ' CONFERMO LE MODIFICHE NEL DB OGNI 100 REGISTRAZIONI
        ' -----------------------------------------------------
        
        If (Indice Mod 100 = 0) Or (Indice = CDec(NVL(TXT_NUMEROMOVIMENTI.Text, 0))) Then
            '
            ' Registro in database
            '
    Timer_ini = Timer
            Pcls_PrimaNota.InserisciPrimaNota StatoFinale
    Timer_fin = Timer
    Timer_diff_registramodifiche = Timer_diff_registramodifiche + (Timer_fin - Timer_ini)
            
            If Pcls_PrimaNota.Stato <> tsOK Then
                MsgBox Pcls_PrimaNota.Errore & " in Pcls_PrimaNota.InserisciPrimaNota"
                Exit Sub
            End If
            
            If Pcls_PrimaNota.PGestRegPN.StatoNonBloccante <> tsNBOk Then
                LST_AVANZAMENTO.AddItem "ERRORE NON BLOCCANTE: " & Pcls_PrimaNota.PGestRegPN.StatoNonBloccante & " - " & Pcls_PrimaNota.PGestRegPN.ErroreNonBloccante
                LST_AVANZAMENTO.ListIndex = LST_AVANZAMENTO.ListCount - 1
            End If
            
            '-------------------------------------------
            ' SCRITTURA DATI EC/PORT
            '-------------------------------------------
            PClsScadenze.RecSetUpdateBatch Connessione
        End If
        
        Set RecSetScadenze = Nothing
    Next
    
    Timer_all_fin = Timer
    Timer_all_diff = Timer_all_fin - Timer_all_ini
    LBL_TEMPO.Caption = Timer_all_diff
    
TXT_TIMER_REG.AddItem Timer_diff_ins_reg
TXT_TIMER_RIGA.AddItem Timer_diff_ins_riga
TXT_TIMER_REGMOD.AddItem Timer_diff_registramodifiche
    
    MsgBox "Le registrazioni sono state inserite"
    
Exit Sub
Err_CMD_INSERT_Click:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub CalcolaImposta(Imponibile As Variant, _
                           CodiceAliquota As Variant, _
                           ByRef Imposta As Variant, _
                           ByRef ImpostaND As Variant)
    
    On Error GoTo Err_CalcolaImposta
    
    If Pobj_PNota Is Nothing Then
        Set Pobj_PNota = New CLSCG_PNOTACHECK
    End If
    
    Pobj_PNota.CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    Pobj_PNota.Valuta = "EURO"
    Pobj_PNota.SommaImponibiliPercIva = Imponibile
    Pobj_PNota.SommaImpostePercIva = 0
    Pobj_PNota.CodiceAliquota = CodiceAliquota
    Pobj_PNota.IndicatoreTipoRegistro = GetTipoRegistro(TXT_CAUSALE.Text)
    Pobj_PNota.TipoCalcoloImposta = CalcoloNormale
    Pobj_PNota.IndicatoreProRata = 0 ' Non gestita
    Pobj_PNota.IndicatoreDetrIva = 2 ' Distinta dal costo
    
    Set Pobj_PNota.GConnect = Connessione
    Pobj_PNota.Sconnect = StrConnect
    
    Pobj_PNota.CalcolaImposta
    
    Imposta = NVL(Pobj_PNota.Imposta, 0)
    ImpostaND = NVL(Pobj_PNota.ImpostaND, 0)
    
Exit Sub
Err_CalcolaImposta:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub AggiornaImposte()
    Dim Imposta1        As Variant
    Dim ImpostaND1      As Variant
    
    On Error Resume Next
    
    If NVL(TXT_IMPONIBILE1.Text, "") <> "" And NVL(TXT_ALIQUOTA1.Text, "") <> "" Then
        CalcolaImposta TXT_IMPONIBILE1.Text, TXT_ALIQUOTA1.Text, Imposta1, ImpostaND1
    Else
        Imposta1 = 0
        ImpostaND1 = 0
    End If
    
    TXT_IMPOSTA1.Text = Imposta1
    TXT_IMPORTOIVA.Text = NVL(Imposta1, 0)
    
    Err.Clear
End Sub

Private Sub TXT_IMPONIBILE1_Validate(Cancel As Boolean)
    On Error Resume Next
    AggiornaImposte
    Err.Clear
End Sub

Private Function GetCondizionePagamento(TipoCliFor As Variant, CodiceCliFor As Variant) As Variant
    Dim Sql     As Variant
    Dim RecSet  As ADODB.Recordset
    
    On Error GoTo Err_GetCondizionePagamento
    
    Sql = "SELECT CG44_CODPAG_CG62" & _
         " FROM CG44_CLIFOR" & _
         " WHERE CG44_DITTA_CG18 = " & CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta & _
         " AND CG44_TIPOCF = " & TipoCliFor & _
         " AND CG44_CLIFOR = " & CodiceCliFor
    Set RecSet = Connessione.Execute(Sql, , adCmdText)
    
    If RecSet.RecordCount > 0 Then
        GetCondizionePagamento = RecSet.Fields("CG44_CODPAG_CG62").Value
    Else
        MsgBox "Errore in GetCondizionePagamento: cliente/fornitore inesistente", vbCritical
    End If
    
Exit Function
Err_GetCondizionePagamento:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Function

Private Function GeneraRecordsetScadenze() As ADODB.Recordset
    Dim SommaIvaEuro            As Variant
    Dim SommaIvaValuta          As Variant
    Dim SommaIvaNonDetEuro      As Variant
    Dim SommaIvaNonDetValuta    As Variant
    Dim RecSetCG43              As ADODB.Recordset
    Dim EsitoCreazioneEcPortOK  As Boolean
    Dim RecSetCG41              As ADODB.Recordset
    
    On Error GoTo Err_GeneraRecordsetScadenze
    
    '
    ' Determino il recordset della testata di prima nota
    '
    Set RecSetCG41 = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata
    RecSetCG41.Filter = "CG41_NUMREG = '" & Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione & "'"
    
    '
    ' Determino la somma dell'IVA detraibile/non detraibile in EURO e in valuta
    '
    SommaIvaEuro = 0
    SommaIvaValuta = 0
    SommaIvaNonDetEuro = 0
    SommaIvaNonDetValuta = 0
    If Not Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetCG43 Is Nothing Then
        Set RecSetCG43 = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetCG43.Clone
        RecSetCG43.Filter = "CG43_NUMREG_CG41 = '" & Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione & "'"
        If RecSetCG43.RecordCount > 0 Then
            RecSetCG43.MoveFirst
            While Not RecSetCG43.EOF
                SommaIvaEuro = SommaIvaEuro + NVL(RecSetCG43.Fields("CG43_IMPOSTA").Value, 0)
                SommaIvaValuta = SommaIvaValuta + NVL(RecSetCG43.Fields("CG43_IMPOSTAVAL").Value, 0)
                SommaIvaNonDetEuro = SommaIvaNonDetEuro + NVL(RecSetCG43.Fields("CG43_IMPOSTAND").Value, 0)
                SommaIvaNonDetValuta = SommaIvaNonDetValuta + NVL(RecSetCG43.Fields("CG43_IMPOSTANDVAL").Value, 0)
                RecSetCG43.MoveNext
            Wend
        End If
    End If
    
    Set PClsScadenze.ClsDittaCorrente = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente
    PClsScadenze.ImportoInEuro = RecSetCG41.Fields("CG41_IMPTOTALE").Value
    PClsScadenze.ImportoInValuta = RecSetCG41.Fields("CG41_IMPTOTALEVAL").Value
    PClsScadenze.IvaInEuro = SommaIvaEuro + SommaIvaNonDetEuro
    PClsScadenze.IvaInValuta = SommaIvaValuta + SommaIvaNonDetValuta
    PClsScadenze.DataDocumento = NVL(RecSetCG41.Fields("CG41_DATADOC").Value, _
                                     RecSetCG41.Fields("CG41_DATAREG").Value)
    PClsScadenze.Valuta = NVL(RecSetCG41.Fields("CG41_CODICE_CG08").Value, "EURO")
    PClsScadenze.Cambio = NVL(RecSetCG41.Fields("CG41_CAMBIO").Value, 0)
    
    PClsScadenze.Conto = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroClienti ' Fornitori ' Clienti
    PClsScadenze.TipoCliFor = NVL(RecSetCG41.Fields("CG41_TIPOCF_CG44").Value, 0)
    PClsScadenze.CodiceCliFor = RecSetCG41.Fields("CG41_CLIFOR_CG44").Value
    PClsScadenze.CondPagamento = GetCondizionePagamento(PClsScadenze.TipoCliFor, PClsScadenze.CodiceCliFor)
    
    PClsScadenze.DareAvere = 1 ' Cliente -> Dare (= 1)
    
    PClsScadenze.TipoElaborazioneScadenze = Inizializzazione
    
    PClsScadenze.Ditta = TXT_DITTA.Text
    PClsScadenze.NumeroRigaContabile = 1 ' riferimento al numero riga contabile che apre la partita
    
    PClsScadenze.NumeroRegistrazione = RecSetCG41.Fields("CG41_NUMREG").Value
    PClsScadenze.CausaleContabile = RecSetCG41.Fields("CG41_CODICE_CG33").Value
    PClsScadenze.NumeroPartita = RecSetCG41.Fields("CG41_NUMDOC").Value
    PClsScadenze.SezionalePartita = RecSetCG41.Fields("CG41_SEZIONALE").Value
    PClsScadenze.PartitaBis = 0
    PClsScadenze.DataRegistrazione = RecSetCG41.Fields("CG41_DATAREG").Value
    PClsScadenze.NumeroDocumento = RecSetCG41.Fields("CG41_NUMDOC").Value
    PClsScadenze.Sezionale = RecSetCG41.Fields("CG41_SEZIONALE").Value
    PClsScadenze.DocumentoBis = RecSetCG41.Fields("CG41_FLGDOCBIS").Value
    PClsScadenze.NumeroDocumentoOrigine = RecSetCG41.Fields("CG41_NUMDOCORIG").Value
    PClsScadenze.FlagAcconto = 0
    PClsScadenze.DescrCausaleContabile = RecSetCG41.Fields("CG41_DESCAUSALE").Value
    PClsScadenze.IndTipoMov = RecSetCG41.Fields("CG41_INDTIPOMOV").Value
    
    EsitoCreazioneEcPortOK = PClsScadenze.GeneraRecordsetScadenze(Connessione)
    
    If Not EsitoCreazioneEcPortOK Then
        MsgBox "Errore in generazione Estratto conto / Portafoglio", vbCritical
    Else
        Set GeneraRecordsetScadenze = PClsScadenze.RecordsetScadenze
    End If
    
    Set RecSetCG41 = Nothing
    Set RecSetCG43 = Nothing
    
Exit Function
Err_GeneraRecordsetScadenze:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Function
