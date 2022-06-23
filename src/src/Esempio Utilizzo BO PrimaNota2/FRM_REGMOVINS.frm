VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_COMBOBOX.ocx"
Object = "{D8EB97B9-26FF-11D2-A1C0-0060082875F9}#6.16#0"; "TMS_EDITDEFCONTO.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.9#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.8#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.7#0"; "TMS_EDITDATE.ocx"
Begin VB.Form FRM_REGMOVINS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inserimento prima nota - diversi a diversi"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "FRM_REGMOVINS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMD_VIEW 
      Caption         =   "Visualizza"
      Height          =   525
      Left            =   8880
      Picture         =   "FRM_REGMOVINS.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6480
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dati relativi alla testata:"
      Height          =   1365
      Left            =   60
      TabIndex        =   26
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
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATARATEOPREC 
         Height          =   300
         Left            =   8130
         TabIndex        =   5
         Top             =   390
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin VB.Label Label18 
         Caption         =   "Data reg. rateo anno prec."
         Height          =   225
         Left            =   8130
         TabIndex        =   44
         Top             =   150
         Width           =   1905
      End
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATARATEOCORR 
         Height          =   300
         Left            =   8130
         TabIndex        =   6
         Top             =   990
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin VB.Label Label16 
         Caption         =   "Data reg. rateo anno corr."
         Height          =   225
         Left            =   8130
         TabIndex        =   43
         Top             =   750
         Width           =   1935
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
         TabIndex        =   37
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Totale registrazione"
         Height          =   225
         Left            =   2880
         TabIndex        =   30
         Top             =   660
         Width           =   1605
      End
      Begin VB.Label Label3 
         Caption         =   "Causale"
         Height          =   225
         Left            =   90
         TabIndex        =   29
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Data reg."
         Height          =   225
         Left            =   2880
         TabIndex        =   28
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Codice ditta"
         Height          =   225
         Left            =   90
         TabIndex        =   27
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.CommandButton CMD_INSERT 
      Caption         =   "Esegui"
      Height          =   525
      Left            =   8880
      Picture         =   "FRM_REGMOVINS.frx":28EC
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4110
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dati relativi al dettaglio:"
      Height          =   2565
      Left            =   60
      TabIndex        =   21
      Top             =   1470
      Width           =   10125
      Begin VB.Frame Frame3 
         Height          =   90
         Left            =   30
         TabIndex        =   22
         Top             =   1290
         Width           =   10035
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1 
         Height          =   300
         Left            =   1020
         TabIndex        =   7
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
         TabIndex        =   13
         Top             =   1440
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin VB.Label Label15 
         Caption         =   "Data fine comp."
         Height          =   225
         Left            =   8130
         TabIndex        =   42
         Top             =   750
         Width           =   1335
      End
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAFINCOMP1 
         Height          =   300
         Left            =   8130
         TabIndex        =   12
         Top             =   990
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin VB.Label Label14 
         Caption         =   "Data inizio comp."
         Height          =   225
         Left            =   8130
         TabIndex        =   41
         Top             =   150
         Width           =   1335
      End
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAINICOMP1 
         Height          =   300
         Left            =   8130
         TabIndex        =   11
         Top             =   390
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO2 
         Height          =   315
         Left            =   3630
         TabIndex        =   17
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
         TabIndex        =   10
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
         TabIndex        =   39
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
         TabIndex        =   14
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
         TabIndex        =   16
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
         TabIndex        =   9
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
         TabIndex        =   15
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
         TabIndex        =   8
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
         TabIndex        =   36
         Top             =   1890
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3000
         TabIndex        =   35
         Top             =   2220
         Width           =   705
      End
      Begin VB.Label Label10 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   34
         Top             =   2220
         Width           =   675
      End
      Begin VB.Label Label7 
         Caption         =   "Conto"
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   32
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Conto"
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label17 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3000
         TabIndex        =   23
         Top             =   990
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Avanzamento:"
      Height          =   1695
      Left            =   60
      TabIndex        =   19
      Top             =   4710
      Width           =   10095
      Begin VB.ListBox LST_AVANZAMENTO 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   90
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   9945
      End
   End
   Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
      Height          =   300
      Left            =   60
      TabIndex        =   38
      Top             =   4380
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
      TabIndex        =   31
      Top             =   4140
      Width           =   2745
   End
End
Attribute VB_Name = "FRM_REGMOVINS"
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
'    Dim clsMovCoinSint  As CIBO_PSMOVCOIN.CLSCI_MOVCI
    
    Dim StatoFinale As StatoFinaleEnum
    
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
    ' Istanzio il b.o. per i movimenti di contabilità analitica
    '
'    Set clsMovCoinSint = New CIBO_PSMOVCOIN.CLSCI_MOVCI
'    Set clsMovCoinSint.CInput.ActiveInterface = CallingForm.ActiveInterface
'    Set clsMovCoinSint.CInput.Connection = Connessione
'    clsMovCoinSint.CInput.Ditta = TXT_DITTA.Text
    
    '
    ' Valorizzo le proprietà della classe che gestisce la prima nota
    '
    
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
        
        '
        ' Date registrazioni ratei
        '
        .DataRateoAnnoPrecedente = NVL(TXT_DATARATEOPREC.Text, Null)
        .DataRateoAnnoCorrente = NVL(TXT_DATARATEOCORR.Text, Null)
    End With
    
    Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
    Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione
    Pcls_PrimaNota.PGestRegPN.InserisciRegistrazione
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRegistrazione"
        Exit Sub
    End If
    
    TXT_NUMREG.Text = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
    
    '
    ' Inserimento testata analitica
    '
'    clsMovCoinSint.NuovaRegistrazione CStr(Date), Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
'
'    clsMovCoinSint.rstRegTestata.Fields("CI12_CODCAUS_CI03").Value = "01"
'    clsMovCoinSint.rstRegTestata.Fields("CI12_DESCRCAUS").Value = "descrizione causale analitica"
'    clsMovCoinSint.rstRegTestata.Fields("CI12_TIPODOC").Value = 66      'Documento CoInd
'    clsMovCoinSint.rstRegTestata.Fields("CI12_INDTIPOMOV").Value = 1    'Da consoldare
'    clsMovCoinSint.rstRegTestata.Fields("CI12_FLGMOVVAR").Value = 1     'Variato
'    clsMovCoinSint.rstRegTestata.Fields("CI12_SEZDOC").Value = "00"     'Sezionale
'    clsMovCoinSint.rstRegTestata.Fields("CI12_NUMDOC").Value = 12345    'Numero Documento
'    clsMovCoinSint.rstRegTestata.Fields("CI12_FLGDOCBIS").Value = 0     'FlgBis
    
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
        
        '
        ' Date competenza
        '
        .FlagCompetenza = SiNo.Si
        .IndicatoreRateiRisconti = IndRateoRiscontoEnum.RateoeRisconto
        .DataInizioCompetenza = NVL(TXT_DATAINICOMP1.Text, Null)
        .DataFineCompetenza = NVL(TXT_DATAFINCOMP1.Text, Null)
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRiga"
        Exit Sub
    End If
    
    '
    ' Inserimento riga contabilità analitica
    '
'    clsMovCoinSint.NuovaRigaRegistrazione
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_SEGNO").Value = 0 'Nessuno
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_RIGACONT_CG42").Value = Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.NumeroRigaCont
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_CODPDC_PC01").Value = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParCOIN.CodicePdC
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_CONTO_PC03").Value = "010000000000001"               'Codice conto CdC
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_CODVDS_PC01").Value = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParCOIN.CodiceVdS
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_VDS_PC03").Value = "000000000012345"                 'Codice conto VdS
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_IMPORTO").Value = Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Importo   'Importo
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_DESCRAGG").Value = "descrizione aggiuntiva riga analitica"
'
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_DADATACOMP").Value = "01/01/2009"
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_ADATACOMP").Value = "04/04/2009"
'
'    clsMovCoinSint.ConfermaRigaRegistrazione
    
'''    clsMovCoinSint.ConfermaRegistrazione
    
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
    
        '
        ' Date competenza
        '
        .FlagCompetenza = SiNo.No
        .IndicatoreRateiRisconti = IndRateoRiscontoEnum.Niente
        .DataInizioCompetenza = Null
        .DataFineCompetenza = Null
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRiga"
        Exit Sub
    End If
    
    '
    ' Inserimento riga contabilità analitica
    '
'    clsMovCoinSint.NuovaRigaRegistrazione
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_SEGNO").Value = 0 'Nessuno
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_RIGACONT_CG42").Value = Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.NumeroRigaCont
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_CODPDC_PC01").Value = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParCOIN.CodicePdC
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_CONTO_PC03").Value = "010000000000001"               'Codice conto CdC
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_CODVDS_PC01").Value = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParCOIN.CodiceVdS
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_VDS_PC03").Value = "000000000012345"                 'Codice conto VdS
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_IMPORTO").Value = Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Importo   'Importo
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_DESCRAGG").Value = "descrizione aggiuntiva riga analitica"
'
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_DADATACOMP").Value = "01/01/2009"
'    clsMovCoinSint.rstRegCorpo.Fields("CI13_ADATACOMP").Value = "04/04/2009"
'
'    clsMovCoinSint.ConfermaRigaRegistrazione

'''    clsMovCoinSint.ConfermaRegistrazione
    
''    Pcls_PrimaNota.PGestRegPN.CPInput.ScriviMasterDoc = False
    
    
'------------------------------------------------------------
'    Pcls_PrimaNota.PGestRegPN.RegistraModifiche
'------------------------------------------------------------

    Pcls_PrimaNota.InserisciPrimaNota StatoFinale
    
    
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.RegistraModifiche"
        Exit Sub
    End If
    
    If Pcls_PrimaNota.PGestRegPN.StatoNonBloccante <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.ErroreNonBloccante
    End If
    
'    clsMovCoinSint.ConfermaRegistrazione
    
'    Set clsMovCoinSint.CInput.ActiveInterface = Nothing
'    Set clsMovCoinSint.CInput.Connection = Nothing
'    clsMovCoinSint.Terminate
'    Set clsMovCoinSint = Nothing
    
    MsgBox "La registrazione è stata inserita", vbInformation
    
Exit Sub
Err_CMD_INSERT_Click:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
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
    
    Set Pcls_Decode = New COBO_LOOKUPDECODE.CLSCO_DECODE
    Set Pcls_Decode.ActiveInterface = CallingForm.ActiveInterface
    
    '
    ' Carico il combo del segno
    '
    CBO_SEGNO1.AddItemData "Dare", 1
    CBO_SEGNO1.AddItemData "Avere", 2
    
    CBO_SEGNO2.AddItemData "Dare", 1
    CBO_SEGNO2.AddItemData "Avere", 2
    
    ' 20.02.2012
'    SDB_SIDEBAR.Initialize CallingForm.ActiveInterface
    
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

Private Sub SDB_SIDEBAR_OnClickItemLink(ByVal objLink As PRJFW_SIDEBAR.CLSFW_LINK)
    If objLink.ObjectNameLink = "ITM50002" Then
        Unload Me
    End If
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

