VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_COMBOBOX.ocx"
Object = "{D8EB97B9-26FF-11D2-A1C0-0060082875F9}#6.15#0"; "TMS_EDITDEFCONTO.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.9#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.8#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.7#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.6#0"; "TMS_EDITDATE.ocx"
Begin VB.Form FRM_REGMOVINS_CASSA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inserimento prima nota - diversi a diversi - CASSA"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "FRM_REGMOVINS_CASSA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Dati relativi al dettaglio:"
      Height          =   2565
      Left            =   60
      TabIndex        =   37
      Top             =   4470
      Width           =   10125
      Begin VB.Frame Frame6 
         Height          =   90
         Left            =   30
         TabIndex        =   38
         Top             =   1290
         Width           =   10035
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1_CASSA 
         Height          =   300
         Left            =   1020
         TabIndex        =   39
         Top             =   240
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO2_CASSA 
         Height          =   300
         Left            =   1020
         TabIndex        =   40
         Top             =   1440
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO2_CASSA 
         Height          =   315
         Left            =   3630
         TabIndex        =   41
         Top             =   2190
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO1_CASSA 
         Height          =   315
         Left            =   3630
         TabIndex        =   42
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG2_CASSA 
         Height          =   300
         Left            =   1020
         TabIndex        =   47
         Top             =   1830
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   529
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   40
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG1_CASSA 
         Height          =   300
         Left            =   1020
         TabIndex        =   48
         Top             =   600
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   529
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   40
         CanRequired     =   0   'False
      End
      Begin VB.Label Label22 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3000
         TabIndex        =   56
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   55
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label20 
         Caption         =   "Conto"
         Height          =   225
         Left            =   120
         TabIndex        =   54
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   53
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label18 
         Caption         =   "Conto"
         Height          =   225
         Left            =   120
         TabIndex        =   52
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   51
         Top             =   2220
         Width           =   675
      End
      Begin VB.Label Label15 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3000
         TabIndex        =   50
         Top             =   2220
         Width           =   705
      End
      Begin VB.Label Label14 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   49
         Top             =   1890
         Width           =   1215
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO1_CASSA 
         Height          =   300
         Left            =   1020
         TabIndex        =   46
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO2_CASSA 
         Height          =   300
         Left            =   1020
         TabIndex        =   45
         Top             =   2190
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO2_CASSA 
         Height          =   300
         Left            =   3060
         TabIndex        =   44
         Top             =   1440
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   36
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO1_CASSA 
         Height          =   300
         Left            =   3060
         TabIndex        =   43
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   36
         CanRequired     =   0   'False
      End
   End
   Begin VB.CommandButton CMD_VIEW 
      Caption         =   "Visualizza"
      Height          =   525
      Left            =   8880
      Picture         =   "FRM_REGMOVINS_CASSA.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9480
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dati relativi alla testata:"
      Height          =   1365
      Left            =   60
      TabIndex        =   22
      Top             =   60
      Width           =   10125
      Begin PRJFW_EDITM.TXT_EDITM TXT_CAUSALE 
         Height          =   300
         Left            =   1350
         TabIndex        =   2
         Top             =   600
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
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREG 
         Height          =   300
         Left            =   4410
         TabIndex        =   1
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGGTEST 
         Height          =   300
         Left            =   1350
         TabIndex        =   4
         Top             =   990
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   529
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   40
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_TOTALE 
         Height          =   300
         Left            =   4410
         TabIndex        =   3
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   1350
         TabIndex        =   0
         Top             =   240
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
      Begin VB.Label Label13 
         Caption         =   "Descrizione agg."
         Height          =   225
         Left            =   90
         TabIndex        =   33
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Totale registrazione"
         Height          =   225
         Left            =   2880
         TabIndex        =   26
         Top             =   660
         Width           =   1605
      End
      Begin VB.Label Label3 
         Caption         =   "Causale"
         Height          =   225
         Left            =   90
         TabIndex        =   25
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Data reg."
         Height          =   225
         Left            =   2880
         TabIndex        =   24
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Codice ditta"
         Height          =   225
         Left            =   90
         TabIndex        =   23
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.CommandButton CMD_INSERT 
      Caption         =   "Esegui"
      Height          =   525
      Left            =   8880
      Picture         =   "FRM_REGMOVINS_CASSA.frx":28EC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7110
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dati relativi al dettaglio:"
      Height          =   2565
      Left            =   60
      TabIndex        =   17
      Top             =   1650
      Width           =   10125
      Begin VB.Frame Frame3 
         Height          =   90
         Left            =   30
         TabIndex        =   18
         Top             =   1290
         Width           =   10035
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1 
         Height          =   300
         Left            =   1020
         TabIndex        =   5
         Top             =   240
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO2 
         Height          =   300
         Left            =   1020
         TabIndex        =   9
         Top             =   1440
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO2 
         Height          =   315
         Left            =   3630
         TabIndex        =   13
         Top             =   2190
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO1 
         Height          =   315
         Left            =   3630
         TabIndex        =   8
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO1 
         Height          =   300
         Left            =   3060
         TabIndex        =   35
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   36
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO2 
         Height          =   300
         Left            =   3060
         TabIndex        =   10
         Top             =   1440
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   36
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO2 
         Height          =   300
         Left            =   1020
         TabIndex        =   12
         Top             =   2190
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO1 
         Height          =   300
         Left            =   1020
         TabIndex        =   7
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG2 
         Height          =   300
         Left            =   1020
         TabIndex        =   11
         Top             =   1830
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   529
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   40
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG1 
         Height          =   300
         Left            =   1020
         TabIndex        =   6
         Top             =   600
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   529
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   40
         CanRequired     =   0   'False
      End
      Begin VB.Label Label12 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   32
         Top             =   1890
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3000
         TabIndex        =   31
         Top             =   2220
         Width           =   705
      End
      Begin VB.Label Label10 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   30
         Top             =   2220
         Width           =   675
      End
      Begin VB.Label Label7 
         Caption         =   "Conto"
         Height          =   225
         Left            =   120
         TabIndex        =   29
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   28
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Conto"
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label17 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3000
         TabIndex        =   19
         Top             =   990
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Avanzamento:"
      Height          =   1695
      Left            =   60
      TabIndex        =   15
      Top             =   7710
      Width           =   10095
      Begin VB.ListBox LST_AVANZAMENTO 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   90
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   9945
      End
   End
   Begin VB.Label Label24 
      BackColor       =   &H0080FFFF&
      Caption         =   "DATI CASSA"
      Height          =   195
      Left            =   60
      TabIndex        =   58
      Top             =   4230
      Width           =   10095
   End
   Begin VB.Label Label23 
      BackColor       =   &H0080FFFF&
      Caption         =   "DATI GESTIONALI"
      Height          =   195
      Left            =   90
      TabIndex        =   57
      Top             =   1440
      Width           =   10095
   End
   Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
      Height          =   300
      Left            =   60
      TabIndex        =   34
      Top             =   7380
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
   Begin VB.Label Label5 
      Caption         =   "Ultimo numero registrazione assegnato"
      Height          =   225
      Left            =   90
      TabIndex        =   27
      Top             =   7140
      Width           =   2745
   End
End
Attribute VB_Name = "FRM_REGMOVINS_CASSA"
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
Private Pcls_Decode                     As COBO_LOOKUPDECODE.CLSCO_DECODE
Private ClsMovGen                       As CGUO_MOVCONTABILI.CLSCG_MOVCONTABILI

Private Sub CMD_INSERT_Click()
    On Error GoTo Err_CMD_INSERT_Click
    
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
    ' Valorizzo le proprietà della classe che gestisce la prima nota
    '
    
    '------------------------------------------------------------------------------------------------------
    ' DATI GESTIONALI
    '------------------------------------------------------------------------------------------------------
    
    '
    ' Inserimento testata
    '
    Set Pcls_PrimaNota.ActiveInterface = CallingForm.ActiveInterface
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata
        .CodiceDitta = TXT_DITTA.Text
        .DataRegistrazione = TXT_DATAREG.Text
        .NumeroRegistrazione = ""
        .CodiceCausale = TXT_CAUSALE.Text
        .NumeroDocumento = 0
        .ImportoDocumento = TXT_TOTALE.Text
        .DescrAggiuntiva = TXT_DESCRAGGTEST.Text
    End With
    
    Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
    Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione
    
    '------------------------------------------------------------------------------------------------------
    Pcls_PrimaNota.PGestRegPN.CPInput.TipoScritturaContab = TipoContabIndicatore_Gestionale
    '------------------------------------------------------------------------------------------------------
    
    Pcls_PrimaNota.PGestRegPN.InserisciRegistrazione
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRegistrazione"
        Exit Sub
    End If
    
    TXT_NUMREG.Text = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
    
    '
    ' Inserimento prima riga
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
        .Conto = TXT_CONTO1.Text
        .DescrAggiuntiva = TXT_DESCRAGG1.Text
        .Importo = TXT_IMPORTO1.Text
        .Segno = CBO_SEGNO1.Text
        .IndicatoreTipoOperazione = MovimentiContabili_DiversiADiversi
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRiga"
        Exit Sub
    End If
    
    '
    ' Inserimento seconda riga
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
        .Conto = TXT_CONTO2.Text
        .DescrAggiuntiva = TXT_DESCRAGG2.Text
        .Importo = TXT_IMPORTO2.Text
        .Segno = CBO_SEGNO2.Text
        .IndicatoreTipoOperazione = MovimentiContabili_DiversiADiversi
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRiga"
        Exit Sub
    End If
    
    Pcls_PrimaNota.PGestRegPN.RegistraModifiche
    
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.RegistraModifiche"
        Exit Sub
    End If
    
    If Pcls_PrimaNota.PGestRegPN.StatoNonBloccante <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.ErroreNonBloccante
    End If
    
    '------------------------------------------------------------------------------------------------------
    ' DATI CASSA
    '------------------------------------------------------------------------------------------------------
    
    '
    ' Inserimento testata
    '
    Set Pcls_PrimaNota.ActiveInterface = CallingForm.ActiveInterface
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata
        .CodiceDitta = TXT_DITTA.Text
        .DataRegistrazione = TXT_DATAREG.Text
        '
        ' Non valorizzo il num.reg. in modo da utilizzare quello dei dati gestionali
        '
        ' .NumeroRegistrazione = ""
        .CodiceCausale = TXT_CAUSALE.Text
        .NumeroDocumento = 0
        .ImportoDocumento = TXT_TOTALE.Text
        .DescrAggiuntiva = TXT_DESCRAGGTEST.Text
    End With
    
    Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
    Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione
    
    '------------------------------------------------------------------------------------------------------
    Pcls_PrimaNota.PGestRegPN.CPInput.TipoScritturaContab = TipoContabIndicatore_Cassa
    '------------------------------------------------------------------------------------------------------
    
    Pcls_PrimaNota.PGestRegPN.InserisciRegistrazione
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRegistrazione"
        Exit Sub
    End If
    
    TXT_NUMREG.Text = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
    
    '
    ' Inserimento prima riga
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
        .Conto = TXT_CONTO1_CASSA.Text
        .DescrAggiuntiva = TXT_DESCRAGG1_CASSA.Text
        .Importo = TXT_IMPORTO1_CASSA.Text
        .Segno = CBO_SEGNO1_CASSA.Text
        .IndicatoreTipoOperazione = MovimentiContabili_DiversiADiversi
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRiga"
        Exit Sub
    End If
    
    '
    ' Inserimento seconda riga
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
        .Conto = TXT_CONTO2_CASSA.Text
        .DescrAggiuntiva = TXT_DESCRAGG2_CASSA.Text
        .Importo = TXT_IMPORTO2_CASSA.Text
        .Segno = CBO_SEGNO2_CASSA.Text
        .IndicatoreTipoOperazione = MovimentiContabili_DiversiADiversi
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRiga"
        Exit Sub
    End If
    
    Pcls_PrimaNota.PGestRegPN.RegistraModifiche
    
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.RegistraModifiche"
        Exit Sub
    End If
    
    If Pcls_PrimaNota.PGestRegPN.StatoNonBloccante <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.ErroreNonBloccante
    End If
    
    MsgBox "La registrazione è stata inserita", vbInformation
    
Exit Sub
Err_CMD_INSERT_Click:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub CMD_TEST_PC_Click()

End Sub

Private Sub CMD_VIEW_Click()
    Dim ClsInterface        As Cinterface
    Dim IdClasse            As Variant
    
    On Error GoTo Err_CMD_VIEW_Click
    
    If NVL(TXT_NUMREG.Text, "") = "" Then
        Exit Sub
    End If
    
    Set ClsMovGen = Nothing
    Set ClsMovGen = New CGUO_MOVCONTABILI.CLSCG_MOVCONTABILI
    Set ClsInterface = ClsMovGen
    
    IdClasse = "CGUO_MOVCONTABILI.CLSCG_MOVCONTABILI"
    Set CallingForm.ActiveInterface.ClsGlobal.CallInterface = ClsInterface
    ClsInterface.Caption = "Movimenti contabili generati"
    
    '
    ' Valorizzo le corrispondenti proprietà della classe
    '
    Set ClsMovGen.Connessione = Connessione
    ClsMovGen.StrConnessione = StrConnect
    ClsMovGen.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    ClsMovGen.NumeroRegistrazione = TXT_NUMREG.Text
    ClsMovGen.IsCallingProgramAnActiveXExe = False
    
    CallingForm.ActiveInterface.ClsGlobal.ExecDll False, IdClasse, True, tsInsert, Normale, 1000, 500
    Set CallingForm.ActiveInterface.ClsGlobal.CallInterface = Nothing
    CallingForm.ActiveInterface.ClsGlobal.RemoveCurrentInterface ClsInterface
    
    ClsInterface.CloseForm
    Set ClsInterface = Nothing
    Set ClsMovGen.Connessione = Nothing
    Set ClsMovGen = Nothing
Exit Sub
Err_CMD_VIEW_Click:
    MsgBox Err.Number & " - " & Err.Description, , "CMD_VIEW_Click"
    Exit Sub
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    TXT_DITTA.Text = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    '
    ' Setto le proprietà dei conti
    '
    SettaProprietaConti
    
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
    Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.GruppoOperatore = "TeamSa"
    Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.Operatore = "TeamSa"
    
    Set Pcls_Decode = New COBO_LOOKUPDECODE.CLSCO_DECODE
    Set Pcls_Decode.ActiveInterface = CallingForm.ActiveInterface
    
    '
    ' Carico il combo del segno
    '
    CBO_SEGNO1.AddItemData "Dare", 1
    CBO_SEGNO1.AddItemData "Avere", 2
    
    CBO_SEGNO2.AddItemData "Dare", 1
    CBO_SEGNO2.AddItemData "Avere", 2
    
    CBO_SEGNO1_CASSA.AddItemData "Dare", 1
    CBO_SEGNO1_CASSA.AddItemData "Avere", 2
    
    CBO_SEGNO2_CASSA.AddItemData "Dare", 1
    CBO_SEGNO2_CASSA.AddItemData "Avere", 2
    
Exit Sub
Err_Form_Load:
    MsgBox Err.Number & " - " & Err.Description, , "FORM_LOAD"
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim RSOggetto       As Object
    
    On Error Resume Next
    
    Set Pcls_Decode = Nothing
    
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
    LST_AVANZAMENTO.AddItem DescrOperazione
    LST_AVANZAMENTO.ListIndex = LST_AVANZAMENTO.ListCount - 1
Exit Sub
Err_Pcls_GestRegPN_Avanzamento:
    MsgBox Err.Number & " - " & Err.Description, , "Pcls_PrimaNota_Avanzamento"
    Exit Sub
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

Private Sub TXT_CONTO1_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Decode.CampoDecodifica = TXT_DESCRCONTO1
    Call Pcls_Decode.Conto(CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc, TXT_CONTO1.Text)
    
    str_SQL = Pcls_Decode.StringaSQL
    Arr_Fields = Pcls_Decode.ArrayFields
    Str_Connect = StrConnect
    
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

Private Sub TXT_CONTO2_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Decode.CampoDecodifica = TXT_DESCRCONTO2
    Call Pcls_Decode.Conto(CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc, TXT_CONTO2.Text)
    
    str_SQL = Pcls_Decode.StringaSQL
    Arr_Fields = Pcls_Decode.ArrayFields
    Str_Connect = StrConnect
    
    Err.Clear
End Sub

Private Sub TXT_CONTO2_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CONTO2_StartLookup
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CONTO2_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CONTO2_StartLookup"
    Exit Sub
End Sub

Private Sub SettaProprietaConti()
    On Error GoTo Err_SettaProprietaConti
    
    Set TXT_CONTO1.Connessione = Connessione
    TXT_CONTO1.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTO1.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    Set TXT_CONTO2.Connessione = Connessione
    TXT_CONTO2.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTO2.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    Set TXT_CONTO1_CASSA.Connessione = Connessione
    TXT_CONTO1_CASSA.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTO1_CASSA.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    Set TXT_CONTO2_CASSA.Connessione = Connessione
    TXT_CONTO2_CASSA.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTO2_CASSA.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
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

