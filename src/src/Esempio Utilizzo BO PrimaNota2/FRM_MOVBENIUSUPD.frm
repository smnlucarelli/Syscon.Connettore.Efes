VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.7#0"; "TMS_COMBOBOX.ocx"
Object = "{D8EB97B9-26FF-11D2-A1C0-0060082875F9}#6.12#0"; "TMS_EDITDEFCONTO.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.7#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.6#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.5#0"; "TMS_EDITNUM.ocx"
Begin VB.Form FRM_MOVBENIUSUPD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modifica prima nota - movimento IVA beni usati"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "FRM_MOVBENIUSUPD.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMD_UPDATE 
      Caption         =   "Esegui"
      Height          =   525
      Left            =   6630
      Picture         =   "FRM_MOVBENIUSUPD.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1305
   End
   Begin VB.CommandButton CMD_VIEW 
      Caption         =   "Visualizza"
      Height          =   525
      Left            =   6060
      Picture         =   "FRM_MOVBENIUSUPD.frx":28EC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7470
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   18
      Top             =   60
      Width           =   6015
      Begin VB.Label Label14 
         Caption         =   "Numero di registrazione"
         Height          =   225
         Left            =   2550
         TabIndex        =   20
         Top             =   240
         Width           =   1755
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
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   1170
         TabIndex        =   0
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
      Begin VB.Label Label2 
         Caption         =   "Codice ditta"
         Height          =   225
         Left            =   90
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton CMD_MODIFICA 
      Caption         =   "Interroga"
      Height          =   525
      Left            =   6120
      Picture         =   "FRM_MOVBENIUSUPD.frx":2A36
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Caption         =   "Campi da modificare nel dettaglio:"
      Height          =   3795
      Left            =   60
      TabIndex        =   16
      Top             =   2400
      Width           =   7875
      Begin PRJFW_EDITM.TXT_EDITM TXT_ALIQUOTAOLD 
         Height          =   300
         Left            =   2130
         TabIndex        =   25
         Top             =   1170
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
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTOOLD 
         Height          =   300
         Left            =   2130
         TabIndex        =   26
         Top             =   510
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         Enabled         =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITM.TXT_EDITM TXT_ALIQUOTANEW 
         Height          =   300
         Left            =   4620
         TabIndex        =   40
         Top             =   1170
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
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTONEW 
         Height          =   300
         Left            =   4620
         TabIndex        =   6
         Top             =   510
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTOIVAOLD 
         Height          =   300
         Left            =   2130
         TabIndex        =   44
         Top             =   2700
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         Enabled         =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTOIVANEW 
         Height          =   300
         Left            =   4620
         TabIndex        =   50
         Top             =   2700
         Width           =   1845
         _ExtentX        =   3281
         _ExtentY        =   529
         IsLookup        =   -1  'True
         Enabled         =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNOIVANEW 
         Height          =   315
         Left            =   4620
         TabIndex        =   51
         Top             =   3360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNOIVAOLD 
         Height          =   315
         Left            =   2130
         TabIndex        =   45
         Top             =   3360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNONEW 
         Height          =   315
         Left            =   4620
         TabIndex        =   43
         Top             =   2160
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNOOLD 
         Height          =   315
         Left            =   2130
         TabIndex        =   34
         Top             =   2160
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Enabled         =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOIVANEW 
         Height          =   300
         Left            =   4620
         TabIndex        =   52
         Top             =   3030
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   90
         X2              =   7755
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label Label27 
         Caption         =   "Segno"
         Height          =   225
         Left            =   480
         TabIndex        =   49
         Top             =   3450
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "Conto di controp. IVA"
         Height          =   225
         Left            =   480
         TabIndex        =   48
         Top             =   2760
         Width           =   2025
      End
      Begin VB.Label Label11 
         Caption         =   "Importo"
         Height          =   225
         Left            =   480
         TabIndex        =   47
         Top             =   3090
         Width           =   675
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOIVAOLD 
         Height          =   300
         Left            =   2130
         TabIndex        =   46
         Top             =   3030
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPONIBILENEW 
         Height          =   300
         Left            =   4620
         TabIndex        =   7
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   11
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPOSTANEW 
         Height          =   300
         Left            =   4620
         TabIndex        =   42
         Top             =   1500
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPOSTANDNEW 
         Height          =   300
         Left            =   4620
         TabIndex        =   41
         Top             =   1830
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label6 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori correnti"
         Height          =   225
         Left            =   2130
         TabIndex        =   39
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori modificati"
         Height          =   225
         Left            =   4620
         TabIndex        =   38
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label Label8 
         Caption         =   "Conto di contropartita"
         Height          =   225
         Left            =   480
         TabIndex        =   37
         Top             =   570
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Cod. aliquota"
         Height          =   225
         Left            =   480
         TabIndex        =   35
         Top             =   1230
         Width           =   1005
      End
      Begin VB.Label Label17 
         Caption         =   "Segno"
         Height          =   225
         Left            =   480
         TabIndex        =   33
         Top             =   2220
         Width           =   585
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPONIBILEOLD 
         Height          =   300
         Left            =   2130
         TabIndex        =   32
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
      Begin VB.Label Label21 
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
         Left            =   90
         TabIndex        =   31
         Top             =   900
         Width           =   315
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPOSTAOLD 
         Height          =   300
         Left            =   2130
         TabIndex        =   30
         Top             =   1500
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label23 
         Caption         =   "Imposta"
         Height          =   225
         Left            =   480
         TabIndex        =   29
         Top             =   1560
         Width           =   735
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPOSTANDOLD 
         Height          =   300
         Left            =   2130
         TabIndex        =   28
         Top             =   1830
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         Enabled         =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label24 
         Caption         =   "Imposta ND"
         Height          =   225
         Left            =   480
         TabIndex        =   27
         Top             =   1890
         Width           =   915
      End
      Begin VB.Label Label9 
         Caption         =   "Imponibile"
         Height          =   225
         Left            =   480
         TabIndex        =   36
         Top             =   900
         Width           =   795
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Avanzamento:"
      Height          =   1695
      Left            =   60
      TabIndex        =   14
      Top             =   6300
      Width           =   5955
      Begin VB.ListBox LST_AVANZAMENTO 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   90
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   5805
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Campi da modificare nella testata:"
      Height          =   1545
      Left            =   60
      TabIndex        =   12
      Top             =   780
      Width           =   7875
      Begin PRJFW_EDIT.TxtEdit TXT_TOTALENEW 
         Height          =   300
         Left            =   4710
         TabIndex        =   5
         Top             =   1170
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGGTESTNEW 
         Height          =   300
         Left            =   4710
         TabIndex        =   3
         Top             =   510
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   529
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   25
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMDOCORIGNEW 
         Height          =   300
         Left            =   4710
         TabIndex        =   4
         Top             =   840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin VB.Label Label5 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori correnti"
         Height          =   225
         Left            =   1470
         TabIndex        =   24
         Top             =   240
         Width           =   3075
      End
      Begin VB.Label Label7 
         BackColor       =   &H00BDFDFC&
         Caption         =   " Valori modificati"
         Height          =   225
         Left            =   4710
         TabIndex        =   23
         Top             =   240
         Width           =   3075
      End
      Begin PRJFW_EDIT.TxtEdit TXT_TOTALEOLD 
         Height          =   300
         Left            =   1470
         TabIndex        =   21
         Top             =   1170
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Enabled         =   0   'False
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMDOCORIGOLD 
         Height          =   300
         Left            =   1470
         TabIndex        =   11
         Top             =   840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Enabled         =   0   'False
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGGTESTOLD 
         Height          =   300
         Left            =   1470
         TabIndex        =   10
         Top             =   510
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   240
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   25
         CanRequired     =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Num. doc. orig."
         Height          =   225
         Left            =   180
         TabIndex        =   17
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Descrizione agg."
         Height          =   225
         Left            =   180
         TabIndex        =   13
         Top             =   540
         Width           =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "Totale"
         Height          =   225
         Left            =   180
         TabIndex        =   22
         Top             =   1200
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FRM_MOVBENIUSUPD"
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
Private Pobj_PNota                      As CLSCG_PNOTACHECK
Private ClsMovGen                       As CGUO_MOVCONTABILI.CLSCG_MOVCONTABILI
Private PClsScadenze                    As PFUO_SCADENZE.CLSPF_SCADENZE

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
    
    '
    ' Visualizzo i dati relativi alla registrazione
    '
    TXT_DESCRAGGTESTOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DESCRAGG").Value
    TXT_DESCRAGGTESTNEW.Text = TXT_DESCRAGGTESTOLD.Text
    TXT_NUMDOCORIGOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMDOCORIG").Value
    TXT_NUMDOCORIGNEW.Text = TXT_NUMDOCORIGOLD.Text
    TXT_TOTALEOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_IMPTOTALE").Value
    TXT_TOTALENEW.Text = TXT_TOTALEOLD.Text
    
    If Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.RecordCount > 0 Then
        Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Filter = "CG42_NUMRIGACONT = 2"
        If Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.RecordCount > 0 Then
            Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.MoveFirst
            TXT_CONTOOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_CONTOPAR_CG24").Value
            TXT_CONTONEW.Text = TXT_CONTOOLD.Text
            TXT_IMPONIBILEOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_IMPTOTALE").Value
            TXT_IMPONIBILENEW.Text = TXT_IMPONIBILEOLD.Text
            CBO_SEGNOOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_INDDAAV").Value
            CBO_SEGNONEW.Text = CBO_SEGNOOLD.Text
            
            Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovIva.Filter = "CG43_NUMRIGAIVA = " & Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_PROGIVA_CG43").Value
            If Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovIva.RecordCount > 0 Then
                TXT_ALIQUOTAOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovIva.Fields("CG43_ALIQIVA1_CG28").Value
                TXT_ALIQUOTANEW.Text = TXT_ALIQUOTAOLD.Text
                TXT_IMPOSTAOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovIva.Fields("CG43_IMPOSTA").Value
                TXT_IMPOSTANEW.Text = TXT_IMPOSTAOLD.Text
                TXT_IMPOSTANDOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovIva.Fields("CG43_IMPOSTAND").Value
                TXT_IMPOSTANDNEW.Text = TXT_IMPOSTANDOLD.Text
            End If
            
            Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Filter = "CG42_INDTIPOOPER = " & DocumentoIVA_IvaSuFatturaCorrispettivo
            
            If Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.RecordCount > 0 Then
                TXT_CONTOIVAOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_CONTOPAR_CG24").Value
                TXT_CONTOIVANEW.Text = TXT_CONTOIVAOLD.Text
                TXT_IMPORTOIVAOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_IMPTOTALE").Value
                TXT_IMPORTOIVANEW.Text = TXT_IMPORTOIVAOLD.Text
                CBO_SEGNOIVAOLD.Text = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Fields("CG42_INDDAAV").Value
                CBO_SEGNOIVANEW.Text = CBO_SEGNOIVAOLD.Text
            End If
        End If
        Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont.Filter = adFilterNone
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
    
    '
    ' Creo la connessione
    '
    Set Connessione = New ADODB.Connection
    Connessione.ConnectionString = StrConnect
    Connessione.CursorLocation = adUseClient
    Connessione.Open
    
    Set CallingForm.ActiveInterface.Connection = Connessione
    
    Set Pcls_PrimaNota = New CGBO_PRIMANOTA.CLSCG_PRIMANOTA
    
    CBO_SEGNOOLD.AddItemData "Dare", 1
    CBO_SEGNOOLD.AddItemData "Avere", 2
    CBO_SEGNONEW.AddItemData "Dare", 1
    CBO_SEGNONEW.AddItemData "Avere", 2
    
    CBO_SEGNOIVAOLD.AddItemData "Dare", 1
    CBO_SEGNOIVAOLD.AddItemData "Avere", 2
    CBO_SEGNOIVANEW.AddItemData "Dare", 1
    CBO_SEGNOIVANEW.AddItemData "Avere", 2
    
Exit Sub
Err_Form_Load:
    MsgBox Err.Number & " - " & Err.Description, , "FORM_LOAD"
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim RSOggetto       As Object
    
    On Error Resume Next
    
    Set Pobj_PNota = Nothing
    
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
    On Error GoTo Err_Pcls_PrimaNota_Avanzamento
    LST_AVANZAMENTO.AddItem DescrOperazione
    LST_AVANZAMENTO.ListIndex = LST_AVANZAMENTO.ListCount - 1
Exit Sub
Err_Pcls_PrimaNota_Avanzamento:
    MsgBox Err.Number & " - " & Err.Description, , "Pcls_PrimaNota_Avanzamento"
    Exit Sub
End Sub

Private Sub TXT_CONTONEW_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CONTONEW_StartLookup
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CONTONEW_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CONTONEW_StartLookup"
    Exit Sub
End Sub

Private Sub SettaProprietaConti()
    On Error GoTo Err_SettaProprietaConti
    
    Set TXT_CONTOOLD.Connessione = Connessione
    TXT_CONTOOLD.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTOOLD.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    Set TXT_CONTONEW.Connessione = Connessione
    TXT_CONTONEW.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTONEW.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    Set TXT_CONTOIVAOLD.Connessione = Connessione
    TXT_CONTOIVAOLD.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTOIVAOLD.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    Set TXT_CONTOIVANEW.Connessione = Connessione
    TXT_CONTOIVANEW.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTOIVANEW.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
Exit Sub
Err_SettaProprietaConti:
    MsgBox Err.Number & " - " & Err.Description, , "SettaProprietaconti"
    Exit Sub
End Sub

Private Function NVL(Valore As Variant, ValIfNull As Variant) As Variant
    On Error GoTo Err_NVL
    
    If IsEmpty(Valore) Or IsNull(Valore) Then
        NVL = ValIfNull
    Else
        If CStr(Valore) = "" Then
            NVL = ValIfNull
        Else
            NVL = Valore
        End If
    End If
Exit Function
Err_NVL:
    NVL = ""
    Err.Clear
    Exit Function
End Function

Private Sub CMD_UPDATE_Click()
    Dim CampiDaModificareTestata        As ModificaCampoTestataEnum
    Dim CampiDaModificareDettaglio      As ModificaCampoDettaglioEnum
    Dim RecSetMovCont                   As ADODB.Recordset
    Dim RecSetScadenze                  As ADODB.Recordset
    Dim AzioneFinale                    As StatoFinaleEnum
    
    On Error GoTo Err_CMD_UPDATE_Click
    
    '
    ' Determino i campi variati della testata
    '
    CampiDaModificareTestata = tsModificaTestataNiente
    CampiDaModificareDettaglio = tsModificaNiente
    
    If NVL(TXT_NUMDOCORIGOLD.Text, "") <> NVL(TXT_NUMDOCORIGNEW.Text, "") Then
        CampiDaModificareTestata = CampiDaModificareTestata + tsModificaTestataNumDocOrigine
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.ImportoDocumento = NVL(TXT_TOTALENEW.Text, 0)
    End If
    
    If NVL(TXT_TOTALEOLD.Text, 0) <> NVL(TXT_TOTALENEW.Text, 0) Then
        CampiDaModificareTestata = CampiDaModificareTestata + tsModificaTestataTotaleDoc
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.ImportoDocumento = NVL(TXT_TOTALENEW.Text, 0)
        
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaImporto
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Imponibile = NVL(TXT_TOTALENEW.Text, 0)
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Importo = NVL(TXT_TOTALENEW.Text, 0)
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
    ' Determino i campi variati del dettaglio (riga 1)
    '
    CampiDaModificareDettaglio = tsModificaNiente
    
    If NVL(TXT_CONTOOLD.Text, "") <> NVL(TXT_CONTONEW.Text, "") Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaConto
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Conto = NVL(TXT_CONTONEW.Text, "")
    End If
    
    If NVL(TXT_IMPONIBILEOLD.Text, 0) <> NVL(TXT_IMPONIBILENEW.Text, 0) Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaImporto
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Imponibile = NVL(TXT_IMPONIBILENEW.Text, 0)
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Importo = NVL(TXT_IMPONIBILENEW.Text, 0)
    End If
    
    If NVL(TXT_IMPOSTAOLD.Text, 0) <> NVL(TXT_IMPOSTANEW.Text, 0) Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaImposta
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Imposta = NVL(TXT_IMPOSTANEW.Text, 0)
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
    ' Determino i campi variati della contropartita IVA
    '
    CampiDaModificareDettaglio = tsModificaNiente
    
    If NVL(TXT_IMPORTOIVAOLD.Text, "") <> NVL(TXT_IMPORTOIVANEW.Text, "") Then
        CampiDaModificareDettaglio = CampiDaModificareDettaglio + tsModificaImporto
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.Importo = NVL(TXT_IMPORTOIVANEW.Text, "")
    End If
    
    If CampiDaModificareDettaglio <> tsModificaNiente Then
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.NumeroRegistrazione = TXT_NUMREG.Text
        
        '
        ' Determino il numero di riga contabile della contropartita iva
        '
        Set RecSetMovCont = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetCG42
        
        RecSetMovCont.Find "CG42_INDTIPOOPER = " & CGBO_TIPI.TipoOperazione.DocumentoIVA_IvaSuFatturaCorrispettivo
        If Not RecSetMovCont.EOF Then
            Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio.NumeroRigaCont = RecSetMovCont.Fields("CG42_NUMRIGACONT").Value
            Pcls_PrimaNota.PGestRegPN.ModificaRiga CampiDaModificareDettaglio
            If Pcls_PrimaNota.PGestRegPN.Stato <> tsOk Then
                MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.ModificaRiga"
                Exit Sub
            End If
        Else
            MsgBox "Errore in determinazione numero riga contabile contropartita iva"
            Exit Sub
        End If
        
        Set RecSetMovCont = Nothing
    End If
    
    '
    ' Registro le modifiche in database
    '
    If CampiDaModificareTestata <> tsModificaTestataNiente Or CampiDaModificareDettaglio <> tsModificaNiente Then
        Pcls_PrimaNota.Status = tsModify
        Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.ScadenzeVariate = True
        Pcls_PrimaNota.CPInput.RegistraEstrattoConto = True
        Pcls_PrimaNota.CPInput.RegistraPortafoglio = True
        
        If FornitoreGestitoARitenute Then
            Pcls_PrimaNota.CPInput.RegistraRitenuteAcconto = True
        Else
            Pcls_PrimaNota.CPInput.RegistraRitenuteAcconto = False
        End If
        
        Pcls_PrimaNota.PGestRegPN.CPInput.GestioneAutomaticaBeniUsati = True
        Pcls_PrimaNota.CPInput.RegistraBeniUsati = True
        
        Set Pcls_PrimaNota.ActiveInterface = CallingForm.ActiveInterface
        
        Pcls_PrimaNota.ModificaPrimaNota AzioneFinale
        
        If Pcls_PrimaNota.Stato <> tsOk Then
            MsgBox Pcls_PrimaNota.Errore & " in PclsPrimaNota.ModificaPrimaNota"
            Exit Sub
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

Private Sub AggiornaImposte()
    Dim Imposta         As Variant
    Dim ImpostaND       As Variant
    
    On Error Resume Next
    
    If NVL(TXT_IMPONIBILENEW.Text, "") <> "" And NVL(TXT_ALIQUOTANEW.Text, "") <> "" Then
        CalcolaImposta TXT_IMPONIBILENEW.Text, TXT_ALIQUOTANEW.Text, Imposta, ImpostaND
    Else
        Imposta = 0
        ImpostaND = 0
    End If
    
    TXT_IMPOSTANEW.Text = Imposta
    TXT_IMPOSTANDNEW.Text = ImpostaND
    
    TXT_IMPORTOIVANEW.Text = NVL(TXT_IMPORTOIVAOLD.Text, 0) - NVL(TXT_IMPOSTAOLD.Text, 0) + NVL(TXT_IMPOSTANEW.Text, 0)
    
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
    Pobj_PNota.IndicatoreTipoRegistro = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_INDTIPOREG").Value
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

Private Sub TXT_IMPONIBILENEW_Validate(Cancel As Boolean)
    On Error Resume Next
    AggiornaImposte
    Err.Clear
End Sub

Private Function GeneraRecordsetScadenze() As ADODB.Recordset
    Dim SommaIvaEuro            As Variant
    Dim SommaIvaValuta          As Variant
    Dim SommaIvaNonDetEuro      As Variant
    Dim SommaIvaNonDetValuta    As Variant
    Dim RecSetCG43              As ADODB.Recordset
    Dim EsitoCreazioneEcPortOK  As Boolean
    Dim PclsInterface           As Cinterface
    
    On Error GoTo Err_GeneraRecordsetScadenze
    
    '
    ' Istanzio la classe per la gestione delle scadenze
    '
    If PClsScadenze Is Nothing Then
        Set PClsScadenze = New PFUO_SCADENZE.CLSPF_SCADENZE
    End If
    Set PclsInterface = PClsScadenze
    
    '
    ' Determino la somma dell'IVA detraibile/non detraibile in EURO e in valuta
    '
    SommaIvaEuro = 0
    SommaIvaValuta = 0
    SommaIvaNonDetEuro = 0
    SommaIvaNonDetValuta = 0
    If Not Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetCG43 Is Nothing Then
        Set RecSetCG43 = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetCG43.Clone
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
    PClsScadenze.ImportoInEuro = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_IMPTOTALE").Value
    PClsScadenze.ImportoInValuta = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_IMPTOTALEVAL").Value
    PClsScadenze.IvaInEuro = SommaIvaEuro + SommaIvaNonDetEuro
    PClsScadenze.IvaInValuta = SommaIvaValuta + SommaIvaNonDetValuta
    PClsScadenze.DataDocumento = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATADOC").Value, _
                                     Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATAREG").Value)
    PClsScadenze.Valuta = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CODICE_CG08").Value, "EURO")
    PClsScadenze.Cambio = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CAMBIO").Value, 0)
    PClsScadenze.TipoCliFor = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_TIPOCF_CG44").Value, 0)
    PClsScadenze.CodiceCliFor = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CLIFOR_CG44").Value
    PClsScadenze.CondPagamento = GetCondizionePagamento(PClsScadenze.TipoCliFor, PClsScadenze.CodiceCliFor)
    
    EsitoCreazioneEcPortOK = PClsScadenze.GeneraRecordsetScadenze(Connessione)
    
    If Not EsitoCreazioneEcPortOK Then
        MsgBox "Errore in generazione Estratto conto / Portafoglio", vbCritical
    Else
        Set GeneraRecordsetScadenze = PClsScadenze.RecordsetScadenze
    End If
    
    PclsInterface.CloseForm
    Set PclsInterface.ClsGlobal = Nothing
    Set PclsInterface.StatusBar = Nothing
    Set PclsInterface.ActiveNavigator = Nothing
    Set PclsInterface = Nothing
    
    Set PClsScadenze.Connessione = Nothing
    Set PClsScadenze = Nothing
    
    Set RecSetCG43 = Nothing
Exit Function
Err_GeneraRecordsetScadenze:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Function

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

Private Function FornitoreGestitoARitenute() As Boolean
    Dim Sql             As Variant
    Dim RecSet          As ADODB.Recordset
    Dim TipoCF          As Variant
    Dim CliFor          As Variant
    
    On Error GoTo Err_FornitoreGestitoARitenute
    
    TipoCF = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_TIPOCF_CG44").Value
    CliFor = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CLIFOR_CG44").Value
    
    Sql = "SELECT CG16_INDSOGGRIT" & _
         " FROM CG44_CLIFOR" & _
         " INNER JOIN CG16_ANAGGEN" & _
         " ON CG16_CODICE = CG44_CODICE_CG16" & _
         " WHERE CG44_DITTA_CG18 = " & CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta & _
         " AND CG44_TIPOCF = " & TipoCF & _
         " AND CG44_CLIFOR = " & CliFor
    Set RecSet = Connessione.Execute(Sql, , adCmdText)
    
    If RecSet.RecordCount > 0 Then
        If NVL(RecSet.Fields("CG16_INDSOGGRIT").Value, 0) <> 0 Then
            FornitoreGestitoARitenute = True
        Else
            FornitoreGestitoARitenute = False
        End If
    Else
        FornitoreGestitoARitenute = False
    End If
    
    Set RecSet = Nothing
    
Exit Function
Err_FornitoreGestitoARitenute:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Function
