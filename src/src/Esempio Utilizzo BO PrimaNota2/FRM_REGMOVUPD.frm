VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.7#0"; "TMS_COMBOBOX.ocx"
Object = "{D8EB97B9-26FF-11D2-A1C0-0060082875F9}#6.11#0"; "TMS_EDITDEFCONTO.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.7#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.6#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.5#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.4#0"; "TMS_EDITDATE.ocx"
Begin VB.Form FRM_REGMOVUPD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modifica prima nota - diversi a diversi"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   Icon            =   "FRM_REGMOVUPD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMD_MODIFICA 
      Caption         =   "Interroga"
      Height          =   525
      Left            =   6120
      Picture         =   "FRM_REGMOVUPD.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1305
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   60
      TabIndex        =   36
      Top             =   0
      Width           =   6015
      Begin VB.Label Label1 
         Caption         =   "Codice ditta"
         Height          =   225
         Left            =   90
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   1170
         TabIndex        =   38
         Top             =   210
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
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
         Height          =   300
         Left            =   4350
         TabIndex        =   1
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
      Begin VB.Label Label14 
         Caption         =   "Numero di registrazione"
         Height          =   225
         Left            =   2550
         TabIndex        =   37
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.CommandButton CMD_VIEW 
      Caption         =   "Visualizza"
      Height          =   525
      Left            =   7740
      Picture         =   "FRM_REGMOVUPD.frx":28EC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6420
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modifica dati relativi alla testata:"
      Height          =   1545
      Left            =   60
      TabIndex        =   29
      Top             =   660
      Width           =   10155
      Begin PRJFW_EDITM.TXT_EDITM TXT_CAUSALEOLD 
         Height          =   300
         Left            =   3570
         TabIndex        =   17
         Top             =   510
         Width           =   1005
         _ExtentX        =   1799
         _ExtentY        =   529
         IsLookup        =   -1  'True
         DisplayFormat   =   "Maiuscolo"
         Enabled         =   0   'False
         MaxChar         =   4
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         NumRighe        =   0
         MaxWidth        =   4
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREGOLD 
         Height          =   300
         Left            =   990
         TabIndex        =   0
         Top             =   510
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         Enabled         =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREGNEW 
         Height          =   300
         Left            =   5310
         TabIndex        =   3
         Top             =   510
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_TOTALENEW 
         Height          =   300
         Left            =   5310
         TabIndex        =   4
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGGTESTNEW 
         Height          =   300
         Left            =   5310
         TabIndex        =   5
         Top             =   1170
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   529
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   30
         CanRequired     =   0   'False
      End
      Begin VB.Label Label7 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori modificati"
         Height          =   225
         Left            =   5310
         TabIndex        =   41
         Top             =   240
         Width           =   3765
      End
      Begin VB.Label Label5 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori correnti"
         Height          =   225
         Left            =   990
         TabIndex        =   40
         Top             =   240
         Width           =   3765
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGGTESTOLD 
         Height          =   300
         Left            =   990
         TabIndex        =   19
         Top             =   1170
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   30
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_TOTALEOLD 
         Height          =   300
         Left            =   990
         TabIndex        =   18
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label13 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   90
         TabIndex        =   34
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Totale reg."
         Height          =   225
         Left            =   90
         TabIndex        =   32
         Top             =   870
         Width           =   1605
      End
      Begin VB.Label Label3 
         Caption         =   "Causale"
         Height          =   225
         Left            =   2880
         TabIndex        =   31
         Top             =   540
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "Data reg."
         Height          =   225
         Left            =   90
         TabIndex        =   30
         Top             =   510
         Width           =   975
      End
   End
   Begin VB.CommandButton CMD_UPDATE 
      Caption         =   "Esegui"
      Height          =   525
      Left            =   8910
      Picture         =   "FRM_REGMOVUPD.frx":2A36
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5250
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modifica dati relativi al dettaglio:"
      Height          =   2955
      Left            =   60
      TabIndex        =   25
      Top             =   2250
      Width           =   10155
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1OLD 
         Height          =   300
         Left            =   960
         TabIndex        =   20
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1NEW 
         Height          =   300
         Left            =   5310
         TabIndex        =   6
         Top             =   480
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO2OLD 
         Height          =   300
         Left            =   960
         TabIndex        =   46
         Top             =   1890
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO2NEW 
         Height          =   300
         Left            =   5310
         TabIndex        =   10
         Top             =   1890
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO2NEW 
         Height          =   315
         Left            =   7350
         TabIndex        =   13
         Top             =   2550
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO2OLD 
         Height          =   315
         Left            =   3840
         TabIndex        =   53
         Top             =   2550
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO1NEW 
         Height          =   315
         Left            =   7320
         TabIndex        =   9
         Top             =   1140
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO1OLD 
         Height          =   315
         Left            =   3840
         TabIndex        =   23
         Top             =   1140
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   105
         X2              =   10005
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label Label22 
         BackColor       =   &H00800000&
         Caption         =   " Riga 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   58
         Top             =   1620
         Width           =   795
      End
      Begin VB.Label Label21 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori modificati"
         Height          =   225
         Left            =   5310
         TabIndex        =   57
         Top             =   1620
         Width           =   3765
      End
      Begin VB.Label Label20 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori correnti"
         Height          =   225
         Left            =   960
         TabIndex        =   56
         Top             =   1620
         Width           =   3765
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG2NEW 
         Height          =   300
         Left            =   5310
         TabIndex        =   11
         Top             =   2220
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   529
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   30
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO2NEW 
         Height          =   300
         Left            =   5310
         TabIndex        =   12
         Top             =   2550
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO2NEW 
         Height          =   300
         Left            =   7230
         TabIndex        =   55
         Top             =   1890
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   23
         CanRequired     =   0   'False
      End
      Begin VB.Label Label16 
         Caption         =   "Conto"
         Height          =   225
         Left            =   120
         TabIndex        =   51
         Top             =   1920
         Width           =   735
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG2OLD 
         Height          =   300
         Left            =   960
         TabIndex        =   49
         Top             =   2220
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   30
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO2OLD 
         Height          =   300
         Left            =   960
         TabIndex        =   48
         Top             =   2550
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO2OLD 
         Height          =   300
         Left            =   2400
         TabIndex        =   47
         Top             =   1890
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   23
         CanRequired     =   0   'False
      End
      Begin VB.Label Label12 
         BackColor       =   &H00800000&
         Caption         =   " Riga 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   45
         Top             =   210
         Width           =   795
      End
      Begin VB.Label Label11 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori modificati"
         Height          =   225
         Left            =   5310
         TabIndex        =   44
         Top             =   210
         Width           =   3765
      End
      Begin VB.Label Label10 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori correnti"
         Height          =   225
         Left            =   960
         TabIndex        =   43
         Top             =   210
         Width           =   3765
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG1NEW 
         Height          =   300
         Left            =   5310
         TabIndex        =   7
         Top             =   810
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   529
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   30
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO1NEW 
         Height          =   300
         Left            =   5310
         TabIndex        =   8
         Top             =   1140
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO1NEW 
         Height          =   300
         Left            =   7230
         TabIndex        =   42
         Top             =   480
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   23
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO1OLD 
         Height          =   300
         Left            =   2400
         TabIndex        =   35
         Top             =   480
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   23
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO1OLD 
         Height          =   300
         Left            =   960
         TabIndex        =   22
         Top             =   1140
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG1OLD 
         Height          =   300
         Left            =   960
         TabIndex        =   21
         Top             =   810
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   30
         CanRequired     =   0   'False
      End
      Begin VB.Label Label4 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   810
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Conto"
         Height          =   225
         Left            =   120
         TabIndex        =   28
         Top             =   510
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label Label17 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3300
         TabIndex        =   26
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Descr. agg."
         Height          =   225
         Left            =   120
         TabIndex        =   50
         Top             =   2220
         Width           =   915
      End
      Begin VB.Label Label18 
         Caption         =   "Importo"
         Height          =   225
         Left            =   120
         TabIndex        =   52
         Top             =   2550
         Width           =   675
      End
      Begin VB.Label Label19 
         Caption         =   "Segno"
         Height          =   225
         Left            =   3300
         TabIndex        =   54
         Top             =   2550
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Avanzamento:"
      Height          =   1695
      Left            =   60
      TabIndex        =   24
      Top             =   5250
      Width           =   7635
      Begin VB.ListBox LST_AVANZAMENTO 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   90
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   7455
      End
   End
End
Attribute VB_Name = "FRM_REGMOVUPD"
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
    
    Pcls_PrimaNota.PGestRegPN.CPInput.SConnect = StrConnect
    Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione
    Pcls_PrimaNota.PGestRegPN.ModificaRegistrazione
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOk Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.ModificaRegistrazione"
        Exit Sub
    End If
    
    '
    ' Valorizzo i campi con i valori correnti della testata e del dettaglio
    '
    TXT_DATAREGOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATAREG").Value
    TXT_CAUSALEOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CODICE_CG33").Value
    TXT_TOTALEOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_IMPTOTALE").Value
    TXT_DESCRAGGTESTOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DESCRAGG").Value
    
    TXT_DATAREGNEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATAREG").Value
    TXT_TOTALENEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_IMPTOTALE").Value
    TXT_DESCRAGGTESTNEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DESCRAGG").Value
    
    If Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.RecordCount > 0 Then
        Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.MoveFirst
        TXT_CONTO1OLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_CONTOPAR_CG24").Value
        TXT_DESCRAGG1OLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_DESCRAGG").Value
        TXT_IMPORTO1OLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_IMPTOTALE").Value
        CBO_SEGNO1OLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_INDDAAV").Value
        TXT_CONTO1NEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_CONTOPAR_CG24").Value
        TXT_DESCRAGG1NEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_DESCRAGG").Value
        TXT_IMPORTO1NEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_IMPTOTALE").Value
        CBO_SEGNO1NEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_INDDAAV").Value
    
        Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.MoveNext
        If Not Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.EOF Then
            TXT_CONTO2OLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_CONTOPAR_CG24").Value
            TXT_DESCRAGG2OLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_DESCRAGG").Value
            TXT_IMPORTO2OLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_IMPTOTALE").Value
            CBO_SEGNO2OLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_INDDAAV").Value
            TXT_CONTO2NEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_CONTOPAR_CG24").Value
            TXT_DESCRAGG2NEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_DESCRAGG").Value
            TXT_IMPORTO2NEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_IMPTOTALE").Value
            CBO_SEGNO2NEW.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_INDDAAV").Value
        End If
    End If
    
    LST_AVANZAMENTO.AddItem String(40, "-")
    LST_AVANZAMENTO.ListIndex = LST_AVANZAMENTO.ListCount - 1
    
Exit Sub
Err_CMD_MODIFICA_Click:
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
    Me.Top = 1200
    
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
    CBO_SEGNO1OLD.AddItemData "Dare", 1
    CBO_SEGNO1OLD.AddItemData "Avere", 2
    CBO_SEGNO1NEW.AddItemData "Dare", 1
    CBO_SEGNO1NEW.AddItemData "Avere", 2
    
    CBO_SEGNO2OLD.AddItemData "Dare", 1
    CBO_SEGNO2OLD.AddItemData "Avere", 2
    CBO_SEGNO2NEW.AddItemData "Dare", 1
    CBO_SEGNO2NEW.AddItemData "Avere", 2
    
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

Private Sub TXT_CONTO1OLD_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Decode.CampoDecodifica = TXT_DESCRCONTO1OLD
    Call Pcls_Decode.Conto(CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdC, TXT_CONTO1OLD.Text)
    
    str_SQL = Pcls_Decode.StringaSQL
    Arr_Fields = Pcls_Decode.ArrayFields
    Str_Connect = StrConnect
    
    Err.Clear
End Sub

Private Sub TXT_CONTO1OLD_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CONTO1OLD_StartLookup
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CONTO1OLD_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CONTO1OLD_StartLookup"
    Exit Sub
End Sub

Private Sub TXT_CONTO1NEW_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Decode.CampoDecodifica = TXT_DESCRCONTO1NEW
    Call Pcls_Decode.Conto(CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdC, TXT_CONTO1NEW.Text)
    
    str_SQL = Pcls_Decode.StringaSQL
    Arr_Fields = Pcls_Decode.ArrayFields
    Str_Connect = StrConnect
    
    Err.Clear
End Sub

Private Sub TXT_CONTO1NEW_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CONTO1NEW_StartLookup
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CONTO1NEW_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CONTO1NEW_StartLookup"
    Exit Sub
End Sub

Private Sub TXT_CONTO2OLD_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Decode.CampoDecodifica = TXT_DESCRCONTO2OLD
    Call Pcls_Decode.Conto(CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdC, TXT_CONTO2OLD.Text)
    
    str_SQL = Pcls_Decode.StringaSQL
    Arr_Fields = Pcls_Decode.ArrayFields
    Str_Connect = StrConnect
    
    Err.Clear
End Sub

Private Sub TXT_CONTO2OLD_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CONTO2OLD_StartLookup
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CONTO2OLD_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CONTO2OLD_StartLookup"
    Exit Sub
End Sub

Private Sub TXT_CONTO2NEW_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Decode.CampoDecodifica = TXT_DESCRCONTO2NEW
    Call Pcls_Decode.Conto(CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdC, TXT_CONTO2NEW.Text)
    
    str_SQL = Pcls_Decode.StringaSQL
    Arr_Fields = Pcls_Decode.ArrayFields
    Str_Connect = StrConnect
    
    Err.Clear
End Sub

Private Sub TXT_CONTO2NEW_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CONTO2NEW_StartLookup
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CONTO2NEW_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CONTO2NEW_StartLookup"
    Exit Sub
End Sub

Private Sub SettaProprietaConti()
    On Error GoTo Err_SettaProprietaConti
    
    Set TXT_CONTO1OLD.Connessione = Connessione
    TXT_CONTO1OLD.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdC
    TXT_CONTO1OLD.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    Set TXT_CONTO1NEW.Connessione = Connessione
    TXT_CONTO1NEW.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdC
    TXT_CONTO1NEW.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    Set TXT_CONTO2OLD.Connessione = Connessione
    TXT_CONTO2OLD.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdC
    TXT_CONTO2OLD.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    Set TXT_CONTO2NEW.Connessione = Connessione
    TXT_CONTO2NEW.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdC
    TXT_CONTO2NEW.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
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

Private Sub CMD_UPDATE_Click()
    Dim CampiDaModificareTestata        As ModificaCampoTestataEnum
    Dim CampiDaModificareDettaglio      As ModificaCampoDettaglioEnum
    
    On Error GoTo Err_CMD_UPDATE_Click
    
    '
    ' Determino i campi variati della testata
    '
    CampiDaModificareTestata = tsModificaTestataNiente
    
    If NVL(TXT_DATAREGOLD.Text, "") <> NVL(TXT_DATAREGNEW.Text, "") Then
        CampiDaModificareTestata = CampiDaModificareTestata + tsModificaTestataDataRegistrazione
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.DataRegistrazione = NVL(TXT_DATAREGNEW.Text, "")
    End If
    
    If NVL(TXT_TOTALEOLD.Text, 0) <> NVL(TXT_TOTALENEW.Text, 0) Then
        CampiDaModificareTestata = CampiDaModificareTestata + tsModificaTestataTotaleDoc
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.ImportoDocumento = NVL(TXT_TOTALENEW.Text, 0)
    End If
    
    If NVL(TXT_DESCRAGGTESTOLD.Text, "") <> NVL(TXT_DESCRAGGTESTNEW.Text, "") Then
        CampiDaModificareTestata = CampiDaModificareTestata + tsModificaTestataDescAggiuntiva
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.DescrAggiuntiva = NVL(TXT_DESCRAGGTESTNEW.Text, "")
    End If
    
    If CampiDaModificareTestata <> tsModificaTestataNiente Then
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione = TXT_NUMREG.Text
        Pcls_PrimaNota.PGestRegPN.ModificaTestata CampiDaModificareTestata
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOk Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.ModificaTestata"
            Exit Sub
        End If
    End If
    
    '
    ' Determino i campi variati del dettaglio (riga 1)
    '
    CampiDaModificareDettaglio = tsModificaNiente
    
    If NVL(TXT_CONTO1OLD.Text, "") <> NVL(TXT_CONTO1NEW.Text, "") Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaConto
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Conto = NVL(TXT_CONTO1NEW.Text, "")
    End If
    
    If NVL(TXT_DESCRAGG1OLD.Text, "") <> NVL(TXT_DESCRAGG1NEW.Text, "") Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaDescrCausaleAgg
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.DescrAggiuntiva = NVL(TXT_DESCRAGG1NEW.Text, "")
    End If
    
    If NVL(TXT_IMPORTO1OLD.Text, 0) <> NVL(TXT_IMPORTO1NEW.Text, 0) Or _
       NVL(CBO_SEGNO1OLD.Text, 0) <> NVL(CBO_SEGNO1NEW.Text, 0) Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaImporto + tsModificaSegno
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Importo = NVL(TXT_IMPORTO1NEW.Text, 0)
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Segno = NVL(CBO_SEGNO1NEW.Text, 0)
    End If
    
    If CampiDaModificareDettaglio <> tsModificaNiente Then
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.NumeroRegistrazione = TXT_NUMREG.Text
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.NumeroRigaCont = 1
        Pcls_PrimaNota.PGestRegPN.ModificaRiga CampiDaModificareDettaglio
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOk Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.ModificaRiga"
            Exit Sub
        End If
    End If
    
    '
    ' Determino i campi variati del dettaglio (riga 2)
    '
    CampiDaModificareDettaglio = tsModificaNiente
    
    If NVL(TXT_CONTO2OLD.Text, "") <> NVL(TXT_CONTO2NEW.Text, "") Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaConto
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Conto = NVL(TXT_CONTO2NEW.Text, "")
    End If
    
    If NVL(TXT_DESCRAGG2OLD.Text, "") <> NVL(TXT_DESCRAGG2NEW.Text, "") Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaDescrCausaleAgg
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.DescrAggiuntiva = NVL(TXT_DESCRAGG2NEW.Text, "")
    End If
    
    If NVL(TXT_IMPORTO2OLD.Text, 0) <> NVL(TXT_IMPORTO2NEW.Text, 0) Or _
       NVL(CBO_SEGNO2OLD.Text, 0) <> NVL(CBO_SEGNO2NEW.Text, 0) Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaImporto + tsModificaSegno
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Importo = NVL(TXT_IMPORTO2NEW.Text, 0)
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Segno = NVL(CBO_SEGNO2NEW.Text, 0)
    End If
    
    If CampiDaModificareDettaglio <> tsModificaNiente Then
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.NumeroRegistrazione = TXT_NUMREG.Text
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.NumeroRigaCont = 2
        Pcls_PrimaNota.PGestRegPN.ModificaRiga CampiDaModificareDettaglio
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOk Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.ModificaRiga"
            Exit Sub
        End If
    End If
    
    '
    ' Registro le modifiche in database
    '
    If CampiDaModificareTestata <> tsModificaTestataNiente Or CampiDaModificareDettaglio <> tsModificaNiente Then
        Pcls_PrimaNota.PGestRegPN.RegistraModifiche
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOk Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.RegistraModifiche"
            Exit Sub
        End If
        
        If Pcls_PrimaNota.PGestRegPN.StatoNonBloccante <> tsOk Then
            MsgBox Pcls_PrimaNota.PGestRegPN.ErroreNonBloccante
        End If
    End If
    
    MsgBox "La registrazione è stata modificata", vbInformation
    
    '
    ' Reinterrogo la registrazione
    '
    CMD_MODIFICA_Click
    
Exit Sub
Err_CMD_UPDATE_Click:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub
