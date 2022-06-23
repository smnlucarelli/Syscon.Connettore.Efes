VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_COMBOBOX.ocx"
Object = "{D8EB97B9-26FF-11D2-A1C0-0060082875F9}#6.17#0"; "TMS_EDITDEFCONTO.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.12#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.9#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.10#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.8#0"; "TMS_EDITDATE.ocx"
Begin VB.Form FRM_REGMOVINS_ECPORT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inserimento prima nota - diversi a diversi con estratto conto / portafoglio"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXT_TEMPITOTALE 
      Height          =   315
      Left            =   750
      TabIndex        =   47
      Top             =   660
      Width           =   2250
   End
   Begin VB.Frame Frame4 
      Caption         =   "Avanzamento:"
      Height          =   1695
      Left            =   30
      TabIndex        =   33
      Top             =   7170
      Width           =   7635
      Begin VB.ListBox LST_AVANZAMENTO 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   90
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dati relativi al dettaglio:"
      Height          =   4005
      Left            =   30
      TabIndex        =   23
      Top             =   2490
      Width           =   7635
      Begin VB.Frame Frame5 
         Height          =   90
         Left            =   30
         TabIndex        =   37
         Top             =   2580
         Width           =   7575
      End
      Begin VB.Frame Frame3 
         Height          =   90
         Left            =   30
         TabIndex        =   24
         Top             =   1290
         Width           =   7575
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1 
         Height          =   300
         Index           =   0
         Left            =   1170
         TabIndex        =   6
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
         Left            =   1170
         TabIndex        =   11
         Top             =   2790
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1 
         Height          =   300
         Index           =   1
         Left            =   1170
         TabIndex        =   38
         Top             =   1470
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO1 
         Height          =   315
         Index           =   1
         Left            =   4200
         TabIndex        =   39
         Top             =   2190
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO1 
         Height          =   300
         Index           =   1
         Left            =   3060
         TabIndex        =   46
         Top             =   1470
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
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO1 
         Height          =   300
         Index           =   1
         Left            =   1170
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
      Begin VB.Label Label18 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   44
         Top             =   1890
         Width           =   915
      End
      Begin VB.Label Label16 
         Caption         =   "Conto cliente"
         Height          =   225
         Left            =   120
         TabIndex        =   43
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Label15 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   42
         Top             =   2220
         Width           =   675
      End
      Begin VB.Label Label14 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3570
         TabIndex        =   41
         Top             =   2220
         Width           =   855
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG1 
         Height          =   300
         Index           =   1
         Left            =   1170
         TabIndex        =   40
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
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO1 
         Height          =   315
         Index           =   0
         Left            =   4200
         TabIndex        =   10
         Top             =   960
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG2 
         Height          =   300
         Left            =   1170
         TabIndex        =   13
         Top             =   3180
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
         Index           =   0
         Left            =   1170
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
      Begin VB.Label Label17 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3570
         TabIndex        =   32
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   31
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "Conto cliente"
         Height          =   225
         Left            =   120
         TabIndex        =   30
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   29
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Conto"
         Height          =   225
         Left            =   120
         TabIndex        =   28
         Top             =   2850
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   3570
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3420
         TabIndex        =   26
         Top             =   3600
         Width           =   705
      End
      Begin VB.Label Label12 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Width           =   1215
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO1 
         Height          =   300
         Index           =   0
         Left            =   1170
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
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO2 
         Height          =   300
         Left            =   1170
         TabIndex        =   14
         Top             =   3540
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO2 
         Height          =   300
         Left            =   3060
         TabIndex        =   12
         Top             =   2790
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
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO1 
         Height          =   300
         Index           =   0
         Left            =   3060
         TabIndex        =   7
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
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO2 
         Height          =   315
         Left            =   4200
         TabIndex        =   15
         Top             =   3540
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
   End
   Begin VB.CommandButton CMD_INSERT 
      Caption         =   "Esegui"
      Height          =   525
      Left            =   6360
      Picture         =   "FRM_REGMOVINS_ECPORT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6570
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dati relativi alla testata:"
      Height          =   1365
      Left            =   30
      TabIndex        =   17
      Top             =   1080
      Width           =   7635
      Begin PRJFW_EDITM.TXT_EDITM TXT_CAUSALE 
         Height          =   300
         Left            =   1350
         TabIndex        =   3
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
      Begin PRJFW_EDITNUM.TxtEditNum TXT_TOTALE 
         Height          =   300
         Left            =   4410
         TabIndex        =   4
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Codice ditta"
         Height          =   225
         Left            =   90
         TabIndex        =   22
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Data reg."
         Height          =   225
         Left            =   2880
         TabIndex        =   21
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Causale"
         Height          =   225
         Left            =   90
         TabIndex        =   20
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Totale registrazione"
         Height          =   225
         Left            =   2880
         TabIndex        =   19
         Top             =   660
         Width           =   1605
      End
      Begin VB.Label Label13 
         Caption         =   "Descrizione agg."
         Height          =   225
         Left            =   90
         TabIndex        =   18
         Top             =   1050
         Width           =   1215
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   1350
         TabIndex        =   1
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
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGGTEST 
         Height          =   300
         Left            =   1350
         TabIndex        =   5
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
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREG 
         Height          =   300
         Left            =   4410
         TabIndex        =   2
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
   End
   Begin VB.CommandButton CMD_VIEW 
      Caption         =   "Visualizza"
      Height          =   525
      Left            =   6360
      Picture         =   "FRM_REGMOVINS_ECPORT.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8940
      Width           =   1305
   End
   Begin PRJFW_EDIT.TxtEdit TXT_NUMEROMOVIMENTI 
      Height          =   300
      Left            =   2490
      TabIndex        =   51
      Top             =   90
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      CanRequired     =   0   'False
   End
   Begin VB.Label Label33 
      Caption         =   "Numero di registrazioni da inserire"
      Height          =   225
      Left            =   0
      TabIndex        =   50
      Top             =   150
      Width           =   2445
   End
   Begin VB.Label Label19 
      Caption         =   "TEMPI:"
      Height          =   195
      Left            =   30
      TabIndex        =   49
      Top             =   750
      Width           =   675
   End
   Begin VB.Label Label22 
      Caption         =   "TOTALE:"
      Height          =   195
      Left            =   780
      TabIndex        =   48
      Top             =   450
      Width           =   1875
   End
   Begin VB.Label Label5 
      Caption         =   "Ultimo numero registrazione assegnato"
      Height          =   225
      Left            =   60
      TabIndex        =   36
      Top             =   6600
      Width           =   2745
   End
   Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
      Height          =   300
      Left            =   30
      TabIndex        =   35
      Top             =   6840
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
End
Attribute VB_Name = "FRM_REGMOVINS_ECPORT"
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
Private PClsScadenze                    As PFUO_SCADENZE.CLSPF_SCADENZE
Private ClsDittaBase                    As CLSCO_DITTABASE

Private Timer_ini_TOT                   As Single
Private Timer_fin_TOT                   As Single

Private Sub CMD_INSERT_Click()
    Dim StatoFinale         As StatoFinaleEnum
    Dim NumMov              As Variant
    Dim Indice              As Variant
    
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
    
    NumMov = NVL(TXT_NUMEROMOVIMENTI.Text, 0)
    
Timer_ini_TOT = Timer

    For Indice = 1 To NumMov
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
        ' Inserimento prima riga cliente
        '
        With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
            .CodiceDitta = TXT_DITTA.Text
            .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
            .Conto = TXT_CONTO1(0).Text
            .DescrAggiuntiva = TXT_DESCRAGG1(0).Text
            .Importo = TXT_IMPORTO1(0).Text
            .Segno = CBO_SEGNO1(0).Text
            .IndicatoreTipoOperazione = MovimentiContabili_DiversiADiversi
        End With
        
        Pcls_PrimaNota.PGestRegPN.InserisciRiga
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRiga"
            Exit Sub
        End If
        
        '
        ' Inserimento seconda riga cliente
        '
        With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
            .CodiceDitta = TXT_DITTA.Text
            .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
            .Conto = TXT_CONTO1(1).Text
            .DescrAggiuntiva = TXT_DESCRAGG1(1).Text
            .Importo = TXT_IMPORTO1(1).Text
            .Segno = CBO_SEGNO1(1).Text
            .IndicatoreTipoOperazione = MovimentiContabili_DiversiADiversi
        End With
        
        Pcls_PrimaNota.PGestRegPN.InserisciRiga
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRiga"
            Exit Sub
        End If
        
        '
        ' Inserimento terza riga
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
        
        '
        ' Genero il recordset delle scadenze per il conto cliente
        ' definito sulla prima riga contabile
        '
        GeneraEScriviRecordsetScadenze TXT_CONTO1(0).TipoCF, Right(NVL(TXT_CONTO1(0).Text, "000000"), 6)
        
        '
        ' Genero il recordset delle scadenze per il conto cliente
        ' definito sulla seconda riga contabile
        '
        GeneraEScriviRecordsetScadenze TXT_CONTO1(1).TipoCF, Right(NVL(TXT_CONTO1(1).Text, "000000"), 6)
        
        '
        ' Genero il recordset delle scadenze per il conto PdC (terza riga)
        '
        
        ' Test Rif. TeamProject: IdAtt. 149691
        
        ' primo record
    '    GeneraEScriviRecordsetScadenze_ContoPdC TXT_CONTO2.Text
        
        ' primo record
    '    GeneraEScriviRecordsetScadenze_ContoPdC TXT_CONTO2.Text
        
    Next
    
Timer_fin_TOT = Timer

    TXT_TEMPITOTALE.Text = Timer_fin_TOT - Timer_ini_TOT
    
    MsgBox "Le registrazioni sono state inserite", vbInformation
    
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
    
    TXT_DATAREG.SetFocus
    
    '
    ' Setto le proprietà dei conti
    '
    SettaProprietaConti
    
    TXT_DATAREG.Text = "12/04/2012"
    TXT_TOTALE.Text = 100
    TXT_CAUSALE.Text = "XAEP"
    
    TXT_CONTO1(0).Text = "1400000001"
    TXT_IMPORTO1(0).Text = 50
    CBO_SEGNO1(0).Text = 1
    
    TXT_CONTO1(1).Text = "1400000004"
    TXT_IMPORTO1(1).Text = 50
    CBO_SEGNO1(1).Text = 1
    
    TXT_CONTO2.Text = "1600010100"
    TXT_IMPORTO2.Text = 100
    CBO_SEGNO2.Text = 2
    
    TXT_NUMEROMOVIMENTI.Text = 1
    
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
    CBO_SEGNO1(0).AddItemData "Dare", 1
    CBO_SEGNO1(0).AddItemData "Avere", 2
    CBO_SEGNO1(1).AddItemData "Dare", 1
    CBO_SEGNO1(1).AddItemData "Avere", 2
    
    CBO_SEGNO2.AddItemData "Dare", 1
    CBO_SEGNO2.AddItemData "Avere", 2
    
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

Private Sub TXT_CONTO1_StartDecode(Index As Integer, Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Decode.CampoDecodificaRagSoc = TXT_DESCRCONTO1
    Call Pcls_Decode.ClienteFornitoreDatiAnagrafici(0, Right(NVL(TXT_CONTO1(Index).Text, "000000"), 6))
    
    str_SQL = Pcls_Decode.StringaSQL
    Arr_Fields = Pcls_Decode.ArrayFields
    Str_Connect = StrConnect
    
    Err.Clear
End Sub

Private Sub TXT_CONTO1_StartLookup(Index As Integer, Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
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
    
    Set TXT_CONTO1(0).Connessione = Connessione
    TXT_CONTO1(0).CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTO1(0).Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    Set TXT_CONTO1(1).Connessione = Connessione
    TXT_CONTO1(1).CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTO1(1).Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
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

'
' TipoCF:
' 0: cliente
' 1: fornitore
'
' CliFor: codice del cliente / fornitore
'
Private Sub GeneraEScriviRecordsetScadenze(TipoCF As Variant, CliFor As Variant)
    Dim EsitoCreazioneEcPortOK  As Boolean
    Dim EsitoUpdateBatchOK      As Boolean
    
    On Error GoTo Err_GeneraRecordsetScadenze
    
    '
    ' Istanzio la classe per la gestione delle scadenze
    '
    If PClsScadenze Is Nothing Then
        Set PClsScadenze = New PFUO_SCADENZE.CLSPF_SCADENZE
    End If
    
    ' 16.06.2006
    CreaClasseDittaBase TXT_DITTA.Text
    ' 16.06.2006
    Set PClsScadenze.ClsDittaCorrente = ClsDittaBase ' CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente
    
    PClsScadenze.ImportoInEuro = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_IMPTOTALE").Value
    PClsScadenze.ImportoInValuta = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_IMPTOTALEVAL").Value
    PClsScadenze.IvaInEuro = 0
    PClsScadenze.IvaInValuta = 0
    PClsScadenze.DataDocumento = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATADOC").Value, _
                                     Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATAREG").Value)
    PClsScadenze.Valuta = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CODICE_CG08").Value, "EURO")
    PClsScadenze.Cambio = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CAMBIO").Value, 0)
    
    PClsScadenze.Conto = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroClienti ' Fornitori ' Clienti
    PClsScadenze.TipoCliFor = TipoCF
    PClsScadenze.CodiceCliFor = CliFor
    PClsScadenze.CondPagamento = GetCondizionePagamento(PClsScadenze.TipoCliFor, PClsScadenze.CodiceCliFor)
    
    PClsScadenze.DareAvere = 2 ' Cliente -> Dare (= 1); Fornitore -> Avere (= 2)
    
    PClsScadenze.TipoElaborazioneScadenze = Inizializzazione
    
    PClsScadenze.Ditta = TXT_DITTA.Text
    PClsScadenze.NumeroRigaContabile = 1 ' riferimento al numero riga contabile che apre la partita
    
    PClsScadenze.NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMREG").Value
    PClsScadenze.CausaleContabile = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CODICE_CG33").Value
    PClsScadenze.NumeroPartita = 46546 ' Numero partita da assegnare
    PClsScadenze.SezionalePartita = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_SEZIONALE").Value
    PClsScadenze.PartitaBis = 0
    PClsScadenze.DataRegistrazione = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATAREG").Value
    PClsScadenze.NumeroDocumento = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMDOC").Value
    PClsScadenze.Sezionale = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_SEZIONALE").Value
    PClsScadenze.DocumentoBis = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_FLGDOCBIS").Value
    PClsScadenze.NumeroDocumentoOrigine = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMDOCORIG").Value
    PClsScadenze.FlagAcconto = 1 ' 1 = Acconto, 0 = No Acconto
    PClsScadenze.DescrCausaleContabile = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DESCAUSALE").Value
    PClsScadenze.IndTipoMov = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_INDTIPOMOV").Value
    
    EsitoCreazioneEcPortOK = PClsScadenze.GeneraRecordsetScadenze(Connessione)
    
    If Not EsitoCreazioneEcPortOK Then
        MsgBox "Errore in generazione Estratto conto / Portafoglio", vbCritical
        Exit Sub
    End If
    
    '
    ' Recordset delle scadenze
    '
    Debug.Print PClsScadenze.RecordsetScadenze.RecordCount
    
    EsitoUpdateBatchOK = PClsScadenze.RecSetUpdateBatch(Connessione)
    If Not EsitoUpdateBatchOK Then
        MsgBox "Errore in scrittura Estratto conto / Portafoglio", vbCritical
        Exit Sub
    End If
    
    PClsScadenze.TerminateClass
    Set PClsScadenze.ClsDittaCorrente = Nothing
    Set PClsScadenze.Connessione = Nothing
    Set PClsScadenze = Nothing
    
Exit Sub
Err_GeneraRecordsetScadenze:
    MsgBox Err.Number & " - " & Err.Description
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

'
' 16.06.2006
' Creo la classe ditta base
'
Private Sub CreaClasseDittaBase(CodDitta As Variant)
    Dim ClsDitta        As CLSCO_DITTE
    
    On Error GoTo Err_CreaClasseDittaBase
    
    Set ClsDitta = New CLSCO_DITTE
    
    Set ClsDitta.ConnectionAdo = Connessione
    Set ClsDittaBase = ClsDitta.GetClasseParametriDitta(CodDitta, tsDittaBase)
    Set ClsDittaBase.ClsDatiGest = ClsDitta.GetClasseParametriDitta(CodDitta, tsDittaGestione)
    Set ClsDittaBase.ClsParCoge = ClsDitta.GetClasseParametriDitta(CodDitta, tsParametriCoge)
    Set ClsDittaBase.ClsParPortafoglio = ClsDitta.GetClasseParametriDitta(CodDitta, tsParametriPortafoglio)
    Set ClsDittaBase.ClsPersPdc = ClsDitta.GetClasseParametriDitta(CodDitta, tsContiPersonalizzati)
    Set ClsDittaBase.ClsDecValute = ClsDitta.GetClasseParametriDitta(CodDitta, tsValute)
Exit Sub
Err_CreaClasseDittaBase:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub GeneraEScriviRecordsetScadenze_ContoPdC(ContoPdC As Variant)
    Dim EsitoCreazioneEcPortOK  As Boolean
    Dim EsitoUpdateBatchOK      As Boolean
    
    On Error GoTo Err_GeneraRecordsetScadenze
    
    '
    ' Istanzio la classe per la gestione delle scadenze
    '
    If PClsScadenze Is Nothing Then
        Set PClsScadenze = New PFUO_SCADENZE.CLSPF_SCADENZE
    End If
    
    CreaClasseDittaBase TXT_DITTA.Text
    
    Set PClsScadenze.ClsDittaCorrente = ClsDittaBase ' CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente
    
    PClsScadenze.ImportoInEuro = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_IMPTOTALE").Value
    PClsScadenze.ImportoInValuta = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_IMPTOTALEVAL").Value
    PClsScadenze.IvaInEuro = 0
    PClsScadenze.IvaInValuta = 0
    PClsScadenze.DataDocumento = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATADOC").Value, _
                                     Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATAREG").Value)
    PClsScadenze.Valuta = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CODICE_CG08").Value, "EURO")
    PClsScadenze.Cambio = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CAMBIO").Value, 0)
    
    PClsScadenze.Conto = ContoPdC
    PClsScadenze.TipoCliFor = 1 ' Null
    PClsScadenze.CodiceCliFor = 0 ' Null
    PClsScadenze.CondPagamento = "331"
    
    PClsScadenze.DareAvere = 0 ' Conto attivo = 1 (dare), conto passivo = 2 (avere)
    
    PClsScadenze.TipoElaborazioneScadenze = Inizializzazione
    
    PClsScadenze.Ditta = TXT_DITTA.Text
    PClsScadenze.NumeroRigaContabile = 3 ' riferimento al numero riga contabile che apre la partita
    
    PClsScadenze.NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMREG").Value
    PClsScadenze.CausaleContabile = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CODICE_CG33").Value
    PClsScadenze.NumeroPartita = 12345 ' Numero partita da assegnare
    PClsScadenze.SezionalePartita = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_SEZIONALE").Value
    PClsScadenze.PartitaBis = 0
    PClsScadenze.DataRegistrazione = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATAREG").Value
    PClsScadenze.NumeroDocumento = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMDOC").Value
    PClsScadenze.Sezionale = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_SEZIONALE").Value
    PClsScadenze.DocumentoBis = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_FLGDOCBIS").Value
    PClsScadenze.NumeroDocumentoOrigine = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMDOCORIG").Value
    PClsScadenze.FlagAcconto = 0 ' 1
    PClsScadenze.DescrCausaleContabile = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DESCAUSALE").Value
    PClsScadenze.IndTipoMov = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_INDTIPOMOV").Value
    
    EsitoCreazioneEcPortOK = PClsScadenze.GeneraRecordsetScadenze(Connessione)
    
    If Not EsitoCreazioneEcPortOK Then
        MsgBox "Errore in generazione Estratto conto / Portafoglio", vbCritical
        Exit Sub
    End If
    
    '
    ' Recordset delle scadenze
    '
    Debug.Print PClsScadenze.RecordsetScadenze.RecordCount
    
    EsitoUpdateBatchOK = PClsScadenze.RecSetUpdateBatch(Connessione)
    If Not EsitoUpdateBatchOK Then
        MsgBox "Errore in scrittura Estratto conto / Portafoglio", vbCritical
        Exit Sub
    End If
    
    PClsScadenze.TerminateClass
    Set PClsScadenze.ClsDittaCorrente = Nothing
    Set PClsScadenze.Connessione = Nothing
    Set PClsScadenze = Nothing
    
Exit Sub
Err_GeneraRecordsetScadenze:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub








''''
''''
''''    Dim EsitoCreazioneEcPortOK  As Boolean
''''    Dim EsitoUpdateBatchOK      As Boolean
''''    Dim ClsDitta                As CLSCO_DITTE
''''    Dim CodiceDitta             As Variant
''''
''''    CodiceDitta = < Codice Azienda >
''''
''''    '
''''    ' Istanzio la classe per la gestione delle scadenze
''''    '
''''    Set PClsScadenze = New PFUO_SCADENZE.CLSPF_SCADENZE
''''
''''    Set ClsDitta = New CLSCO_DITTE
''''
''''    Set ClsDitta.ConnectionAdo = Connessione
''''    Set ClsDittaBase = ClsDitta.GetClasseParametriDitta(CodiceDitta, tsDittaBase)
''''    Set ClsDittaBase.ClsDatiGest = ClsDitta.GetClasseParametriDitta(CodiceDitta, tsDittaGestione)
''''    Set ClsDittaBase.ClsParCoge = ClsDitta.GetClasseParametriDitta(CodiceDitta, tsParametriCoge)
''''    Set ClsDittaBase.ClsParPortafoglio = ClsDitta.GetClasseParametriDitta(CodiceDitta, tsParametriPortafoglio)
''''    Set ClsDittaBase.ClsPersPdc = ClsDitta.GetClasseParametriDitta(CodiceDitta, tsContiPersonalizzati)
''''    Set ClsDittaBase.ClsDecValute = ClsDitta.GetClasseParametriDitta(CodiceDitta, tsValute)
''''
''''    Set PClsScadenze.ClsDittaCorrente = ClsDittaBase
''''
''''    PClsScadenze.ImportoInEuro = < totale documento in EURO >
''''    PClsScadenze.ImportoInValuta = < totale documento in VALUTA >
''''    PClsScadenze.IvaInEuro = < totale IVA in EURO >
''''    PClsScadenze.IvaInValuta = < totale IVA in VALUTA >
''''    PClsScadenze.DataDocumento = < data documento >
''''    PClsScadenze.Valuta = < codice valuta >
''''    PClsScadenze.Cambio = < cambio valuta >
''''
''''    PClsScadenze.Conto = < mastro clienti per fatture di vendita, mastro fornitori per fatture di acquisto, conto PdC per altri movimenti >
''''    PClsScadenze.TipoCliFor = < 0 per clienti, 1 per fornitori, NULL per conti PdC >
''''    PClsScadenze.CodiceCliFor = < codice cli/for clienti / fornitori, NULL per conti PdC >
''''    PClsScadenze.CondPagamento = < condizione di pagamento >
''''
''''    PClsScadenze.DareAvere = < 1 per clienti / conti attivi, 2 per fornitori / conti passivi >
''''
''''    PClsScadenze.TipoElaborazioneScadenze = Inizializzazione
''''
''''    PClsScadenze.Ditta = CodiceDitta
''''    PClsScadenze.NumeroRigaContabile = < riferimento al numero riga contabile che apre la partita >
''''    PClsScadenze.NumeroRegistrazione = < numero registrazione del movimento di prima nota >
''''    PClsScadenze.CausaleContabile = < causale contabile >
''''    PClsScadenze.NumeroPartita = < numero partita da assegnare >
''''    PClsScadenze.SezionalePartita = < sezionale della partita >
''''    PClsScadenze.PartitaBis = < 0 = no bis, 1 = bis >
''''    PClsScadenze.DataRegistrazione = < data registrazione >
''''    PClsScadenze.NumeroDocumento = < numero protocollo della registrazione di prima nota >
''''    PClsScadenze.Sezionale = < sezionale della registrazione di prima nota >
''''    PClsScadenze.DocumentoBis = < bis relativo al protocollo della registrazione di prima nota: 0 = no bis, 1 = bis >
''''    PClsScadenze.NumeroDocumentoOrigine = < numero documento origine >
''''    PClsScadenze.FlagAcconto = < = 1 nel caso di registrazione di un acconto, esempio con causale 351, 0 altrimenti >
''''    PClsScadenze.DescrCausaleContabile = < descrizione causale del movimento di prima nota >
''''    PClsScadenze.IndTipoMov = < tipo movimento prima nota: 0 = consolidato, 1 = da consolidare, 2 = previsionale >
''''
''''    EsitoCreazioneEcPortOK = PClsScadenze.GeneraRecordsetScadenze(< oggetto connessione >)
''''
''''    If Not EsitoCreazioneEcPortOK Then
''''        MsgBox "Errore in generazione Estratto conto / Portafoglio", vbCritical
''''        Exit Sub
''''    End If
''''
''''    '
''''    ' Recordset delle scadenze
''''    '
''''    Debug.Print PClsScadenze.RecordsetScadenze.RecordCount
''''
''''    '
''''    ' Il metodo RecSetUpdateBatch di PFUO_SCADENZE.CLSPF_SCADENZE serve per registrare nel database i dati relativi alla partita e alle scadenze
''''    '
''''    EsitoUpdateBatchOK = PClsScadenze.RecSetUpdateBatch(Connessione)
''''    If Not EsitoUpdateBatchOK Then
''''        MsgBox "Errore in scrittura Estratto conto / Portafoglio", vbCritical
''''        Exit Sub
''''    End If
''''
''''    PClsScadenze.TerminateClass
''''    Set PClsScadenze.ClsDittaCorrente = Nothing
''''    Set PClsScadenze.Connessione = Nothing
''''    Set PClsScadenze = Nothing
''''
''''
''''
''''
