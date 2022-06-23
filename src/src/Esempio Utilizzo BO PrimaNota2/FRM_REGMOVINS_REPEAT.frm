VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.9#0"; "TMS_COMBOBOX.ocx"
Object = "{D8EB97B9-26FF-11D2-A1C0-0060082875F9}#6.14#0"; "TMS_EDITDEFCONTO.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.9#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.8#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.7#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.6#0"; "TMS_EDITDATE.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRM_REGMOVINS_REPEAT 
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   15055
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Inserimento"
      TabPicture(0)   =   "FRM_REGMOVINS_REPEAT.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TXT_NUMREG"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CMD_INSERT"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CMD_VIEW"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Variazione"
      TabPicture(1)   =   "FRM_REGMOVINS_REPEAT.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "CMD_QUERY"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CommandButton CMD_VIEW 
         Caption         =   "Visualizza"
         Height          =   525
         Left            =   8880
         Picture         =   "FRM_REGMOVINS_REPEAT.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   7920
         Width           =   1305
      End
      Begin VB.CommandButton CMD_QUERY 
         Caption         =   "Esegui"
         Height          =   525
         Left            =   -66120
         Picture         =   "FRM_REGMOVINS_REPEAT.frx":0182
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1980
         Width           =   1305
      End
      Begin VB.Frame Frame6 
         Height          =   1605
         Left            =   -74940
         TabIndex        =   60
         Top             =   330
         Width           =   10125
         Begin VB.ListBox TXT_TIMER_QUERY 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   1200
            Left            =   7650
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   360
            Width           =   1785
         End
         Begin VB.CommandButton CMD_CLEAR_QUERY 
            Caption         =   "Cancella"
            Height          =   255
            Left            =   9300
            TabIndex        =   61
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label24 
            Caption         =   "NUMREG da interrogare"
            Height          =   225
            Left            =   90
            TabIndex        =   68
            Top             =   270
            Width           =   1815
         End
         Begin PRJFW_EDIT.TxtEdit TXT_NUMITER_QUERY 
            Height          =   300
            Left            =   1980
            TabIndex        =   67
            Top             =   630
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin PRJFW_EDIT.TxtEdit TXT_NUMREG_QUERY 
            Height          =   300
            Left            =   1980
            TabIndex        =   63
            Top             =   240
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
         Begin VB.Label Label25 
            Caption         =   "modifica registrazione"
            Height          =   225
            Left            =   7650
            TabIndex        =   65
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "Numero di iterazioni"
            Height          =   225
            Left            =   90
            TabIndex        =   64
            Top             =   660
            Width           =   1605
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dati relativi alla testata:"
         Height          =   1365
         Left            =   60
         TabIndex        =   39
         Top             =   1950
         Width           =   10125
         Begin PRJFW_EDITM.TXT_EDITM TXT_CAUSALE 
            Height          =   300
            Left            =   1350
            TabIndex        =   40
            Top             =   600
            Width           =   1005
            _ExtentX        =   1799
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
            TabIndex        =   41
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            IsDbField       =   0   'False
            MaxWidth        =   11
            MaxChar         =   13
            CanRequired     =   0   'False
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATARATEOPREC 
            Height          =   300
            Left            =   8130
            TabIndex        =   53
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
            TabIndex        =   52
            Top             =   150
            Width           =   1905
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATARATEOCORR 
            Height          =   300
            Left            =   8130
            TabIndex        =   51
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
            TabIndex        =   50
            Top             =   750
            Width           =   1935
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREG 
            Height          =   300
            Left            =   4410
            TabIndex        =   49
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
            TabIndex        =   48
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
         Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
            Height          =   300
            Left            =   1350
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   1050
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Totale registrazione"
            Height          =   225
            Left            =   2880
            TabIndex        =   45
            Top             =   660
            Width           =   1605
         End
         Begin VB.Label Label3 
            Caption         =   "Causale"
            Height          =   225
            Left            =   90
            TabIndex        =   44
            Top             =   690
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Data reg."
            Height          =   225
            Left            =   2880
            TabIndex        =   43
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Codice ditta"
            Height          =   225
            Left            =   90
            TabIndex        =   42
            Top             =   330
            Width           =   975
         End
      End
      Begin VB.CommandButton CMD_INSERT 
         Caption         =   "Esegui"
         Height          =   525
         Left            =   8880
         Picture         =   "FRM_REGMOVINS_REPEAT.frx":02CC
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5970
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dati relativi al dettaglio:"
         Height          =   2565
         Left            =   60
         TabIndex        =   14
         Top             =   3360
         Width           =   10125
         Begin VB.Frame Frame3 
            Height          =   90
            Left            =   30
            TabIndex        =   15
            Top             =   1290
            Width           =   10035
         End
         Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1 
            Height          =   300
            Left            =   1020
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   20
            Top             =   960
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            MaxChar         =   8
            IsDbField       =   0   'False
            DbCol           =   0
            CanRequired     =   0   'False
         End
         Begin PRJFW_EDIT.TxtEdit TXT_DESCRAGG2 
            Height          =   300
            Left            =   1020
            TabIndex        =   19
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
            TabIndex        =   21
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
         Begin VB.Label Label15 
            Caption         =   "Data fine comp."
            Height          =   225
            Left            =   8130
            TabIndex        =   37
            Top             =   750
            Width           =   1335
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAFINCOMP1 
            Height          =   300
            Left            =   8130
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   150
            Width           =   1335
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAINICOMP1 
            Height          =   300
            Left            =   8130
            TabIndex        =   34
            Top             =   390
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            IsCalendario    =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin PRJFW_EDIT.TxtEdit TXT_DESCRCONTO1 
            Height          =   300
            Left            =   3060
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            IsDbField       =   0   'False
            MaxWidth        =   11
            MaxChar         =   13
            CanRequired     =   0   'False
         End
         Begin VB.Label Label12 
            Caption         =   "Descr. agg."
            Height          =   225
            Left            =   120
            TabIndex        =   29
            Top             =   1890
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Segno"
            Height          =   225
            Left            =   3000
            TabIndex        =   28
            Top             =   2220
            Width           =   705
         End
         Begin VB.Label Label10 
            Caption         =   "Importo"
            Height          =   225
            Left            =   120
            TabIndex        =   27
            Top             =   2220
            Width           =   675
         End
         Begin VB.Label Label7 
            Caption         =   "Conto"
            Height          =   225
            Left            =   120
            TabIndex        =   26
            Top             =   1500
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Descr. agg."
            Height          =   225
            Left            =   120
            TabIndex        =   25
            Top             =   660
            Width           =   915
         End
         Begin VB.Label Label8 
            Caption         =   "Conto"
            Height          =   225
            Left            =   120
            TabIndex        =   24
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Importo"
            Height          =   225
            Left            =   120
            TabIndex        =   23
            Top             =   990
            Width           =   675
         End
         Begin VB.Label Label17 
            Caption         =   "Segno"
            Height          =   225
            Left            =   3000
            TabIndex        =   22
            Top             =   990
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Avanzamento:"
         Height          =   1305
         Left            =   60
         TabIndex        =   12
         Top             =   6570
         Width           =   10095
         Begin VB.ListBox LST_AVANZAMENTO 
            Appearance      =   0  'Flat
            Height          =   1005
            Left            =   90
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   9945
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1605
         Left            =   60
         TabIndex        =   1
         Top             =   330
         Width           =   10125
         Begin VB.CommandButton CMD_CLEAR 
            Caption         =   "Cancella"
            Height          =   255
            Left            =   9300
            TabIndex        =   59
            Top             =   120
            Width           =   795
         End
         Begin VB.ListBox TXT_TIMER_REGMOD 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   1200
            Left            =   7980
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   360
            Width           =   1785
         End
         Begin VB.ListBox TXT_TIMER_RIGA 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   1200
            Left            =   6150
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   360
            Width           =   1785
         End
         Begin VB.ListBox TXT_TIMER_REG 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   1200
            Left            =   4320
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   360
            Width           =   1785
         End
         Begin VB.TextBox TXT_NUMREG_RM 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2640
            TabIndex        =   2
            Text            =   "100"
            Top             =   1230
            Width           =   705
         End
         Begin VB.OptionButton CHK_REGMOD_UNICO 
            Caption         =   "Invoca 'RegistraModifiche' una sola volta"
            Height          =   255
            Left            =   90
            TabIndex        =   4
            Top             =   960
            Value           =   -1  'True
            Width           =   3315
         End
         Begin VB.OptionButton CHK_REGMOD_OGNITOT 
            Caption         =   "Invoca 'RegistraModifiche' ogni                  registrazioni"
            Height          =   255
            Left            =   90
            TabIndex        =   3
            Top             =   1260
            Width           =   4545
         End
         Begin PRJFW_EDIT.TxtEdit TXT_NUMEROMOVIMENTI 
            Height          =   300
            Left            =   2520
            TabIndex        =   10
            Top             =   150
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
            Left            =   90
            TabIndex        =   11
            Top             =   210
            Width           =   2445
         End
         Begin VB.Label Label19 
            Caption         =   "reg. mod."
            Height          =   225
            Left            =   7980
            TabIndex        =   9
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label20 
            Caption         =   "ins. registrazione"
            Height          =   225
            Left            =   4320
            TabIndex        =   8
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label Label21 
            Caption         =   "ins. riga"
            Height          =   225
            Left            =   6150
            TabIndex        =   7
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label22 
            Caption         =   "Coppie di righe D/A da inserire"
            Height          =   225
            Left            =   90
            TabIndex        =   6
            Top             =   570
            Width           =   2325
         End
         Begin PRJFW_EDIT.TxtEdit TXT_NUM_COPPIE_D_A 
            Height          =   300
            Left            =   2520
            TabIndex        =   5
            Top             =   510
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
         Height          =   300
         Left            =   60
         TabIndex        =   55
         Top             =   6240
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
         TabIndex        =   54
         Top             =   6000
         Width           =   2745
      End
   End
End
Attribute VB_Name = "FRM_REGMOVINS_REPEAT"
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

Private Timer_ini                       As Single
Private Timer_fin                       As Single

Private Timer_diff_ins_reg              As Single
Private Timer_diff_ins_riga             As Single
Private Timer_diff_registramodifiche    As Single
Private Timer_diff_modificareg          As Single

Private Sub CMD_CLEAR_Click()
    On Error Resume Next
    
    TXT_TIMER_REG.Clear
    TXT_TIMER_RIGA.Clear
    TXT_TIMER_REGMOD.Clear
    
    Err.Clear
End Sub

Private Sub CMD_CLEAR_QUERY_Click()
    On Error Resume Next
    
    TXT_TIMER_QUERY.Clear
    
    Err.Clear
End Sub

Private Sub CMD_INSERT_Click()
    Dim NumMov              As Variant
    Dim NumCoppie_DA        As Variant
    Dim Indice              As Variant
    Dim Indice_Righe        As Variant
    
    On Error GoTo Err_CMD_INSERT_Click
    
    If NVL(TXT_NUMEROMOVIMENTI.Text, 0) <= 0 Then
        Exit Sub
    End If
    
    NumMov = NVL(TXT_NUMEROMOVIMENTI.Text, 0)
    NumCoppie_DA = NVL(TXT_NUM_COPPIE_D_A.Text, 0)
    
    '
    ' Pulisco il list box che segnala le operazioni
    '
    LST_AVANZAMENTO.Clear
    LST_AVANZAMENTO.Refresh
    
Timer_diff_ins_reg = 0
Timer_diff_ins_riga = 0
Timer_diff_registramodifiche = 0
    
    For Indice = 1 To NumMov
        If Pcls_PrimaNota Is Nothing Then
            Set Pcls_PrimaNota = New CGBO_PRIMANOTA.CLSCG_PRIMANOTA
        End If
        
        '
        ' Pulisco il text che contiene l'ultimo numero di registrazione assegnato
        '
        TXT_NUMREG.Text = ""
        
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
''            .DataRateoAnnoPrecedente = NVL(TXT_DATARATEOPREC.Text, Null)
''            .DataRateoAnnoCorrente = NVL(TXT_DATARATEOCORR.Text, Null)
        End With
        
        Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
        Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione
        
Timer_ini = Timer
        
        Pcls_PrimaNota.PGestRegPN.InserisciRegistrazione
        
Timer_fin = Timer
Timer_diff_ins_reg = Timer_diff_ins_reg + (Timer_fin - Timer_ini)
        
        
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRegistrazione"
            Exit Sub
        End If
        
        TXT_NUMREG.Text = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
        
        For Indice_Righe = 1 To NumCoppie_DA
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
'                .FlagCompetenza = SiNo.Si
'                .IndicatoreRateiRisconti = IndRateoRiscontoEnum.RateoeRisconto
'                .DataInizioCompetenza = NVL(TXT_DATAINICOMP1.Text, Null)
'                .DataFineCompetenza = NVL(TXT_DATAFINCOMP1.Text, Null)
            End With
            
    Timer_ini = Timer
    ' 25.02.2010
            Pcls_PrimaNota.PGestRegPN.InserisciRiga
            
    Timer_fin = Timer
    Timer_diff_ins_riga = Timer_diff_ins_riga + (Timer_fin - Timer_ini)
            
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
            
                '
                ' Date competenza
                '
'                .FlagCompetenza = SiNo.No
'                .IndicatoreRateiRisconti = IndRateoRiscontoEnum.Niente
'                .DataInizioCompetenza = Null
'                .DataFineCompetenza = Null
            End With
            
    Timer_ini = Timer
            
    ' 25.02.2010
          Pcls_PrimaNota.PGestRegPN.InserisciRiga
            
    Timer_fin = Timer
    Timer_diff_ins_riga = Timer_diff_ins_riga + (Timer_fin - Timer_ini)
            
            If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
                MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.InserisciRiga"
                Exit Sub
            End If
        Next

If Not CHK_REGMOD_UNICO.Value Then
    If Indice Mod TXT_NUMREG_RM.Text = 0 Then

Timer_ini = Timer

''    Pcls_PrimaNota.PGestRegPN.CPInput.ScriviMasterDoc = False

' 24.02.2010
        Pcls_PrimaNota.PGestRegPN.RegistraModifiche
        
Timer_fin = Timer
Timer_diff_registramodifiche = Timer_diff_registramodifiche + (Timer_fin - Timer_ini)
    
    End If
End If

        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.RegistraModifiche"
            Exit Sub
        End If
        
        If Pcls_PrimaNota.PGestRegPN.StatoNonBloccante <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.ErroreNonBloccante
        End If
        
        ''' DISTRUZIONE PARZIALE
        Set Pcls_PrimaNota.PGestRegPN.CPInput.DatiRegCollegate.RsDatiDocumentoPerCogeproge = Nothing
        Set Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMasterDoc = Nothing
        Set Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata = Nothing
        Set Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovCont = Nothing
        Set Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetMovIva = Nothing
        Pcls_PrimaNota.DestroyScadPort
        ''' DISTRUZIONE PARZIALE
    Next
    
If CHK_REGMOD_UNICO.Value Then

Timer_ini = Timer

        Pcls_PrimaNota.PGestRegPN.RegistraModifiche
        
Timer_fin = Timer
Timer_diff_registramodifiche = Timer_diff_registramodifiche + (Timer_fin - Timer_ini)
        
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.RegistraModifiche"
            Exit Sub
        End If
        
        If Pcls_PrimaNota.PGestRegPN.StatoNonBloccante <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.ErroreNonBloccante
        End If
        
End If
    
    
TXT_TIMER_REG.AddItem Timer_diff_ins_reg
TXT_TIMER_RIGA.AddItem Timer_diff_ins_riga
TXT_TIMER_REGMOD.AddItem Timer_diff_registramodifiche
    
    
'' 05.02.2010
'MsgBox "Timer_diff_1 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_1 & vbCrLf & _
'       "Timer_diff_2 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_2 & vbCrLf & _
'       "Timer_diff_3 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_3 & vbCrLf & _
'       "Timer_diff_4 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_4 & vbCrLf & _
'       "Timer_diff_5 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_5 & vbCrLf & _
'       "Timer_diff_6 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_6 & vbCrLf & _
'       "Timer_diff_7 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_7 & vbCrLf & _
'       "Timer_diff_8_1 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_8_1 & vbCrLf & _
'       "Timer_diff_8_2 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_8_2 & vbCrLf & _
'       "Timer_diff_8_3 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_8_3 & vbCrLf & _
'       "Timer_diff_8_4 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_8_4 & vbCrLf & _
'       "Timer_diff_9 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_9 & vbCrLf & _
'       "Timer_diff_10 = " & Pcls_PrimaNota.PGestRegPN.Timer_diff_10
'
'
'Pcls_PrimaNota.PGestRegPN.Timer_diff_1 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_2 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_3 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_4 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_5 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_6 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_7 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_8_1 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_8_2 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_8_3 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_8_4 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_9 = 0
'Pcls_PrimaNota.PGestRegPN.Timer_diff_10 = 0
'

    DistruggiBOPrimaNota
    
    MsgBox "Le registrazioni sono state inserite", vbInformation
    
Exit Sub
Err_CMD_INSERT_Click:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub CMD_QUERY_Click()
    Dim NumIterazioni       As Variant
    Dim Indice              As Variant
    
    On Error GoTo Err_CMD_QUERY_Click
    
    '
    ' Pulisco il list box che segnala le operazioni
    '
    LST_AVANZAMENTO.Clear
    LST_AVANZAMENTO.Refresh
    
    NumIterazioni = NVL(TXT_NUMITER_QUERY.Text, 0)
    
Timer_diff_modificareg = 0
    
    For Indice = 1 To NumIterazioni
        If Pcls_PrimaNota Is Nothing Then
            Set Pcls_PrimaNota = New CGBO_PRIMANOTA.CLSCG_PRIMANOTA
        End If
        
        Set Pcls_PrimaNota.ActiveInterface = CallingForm.ActiveInterface
        With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata
            .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
            .NumeroRegistrazione = TXT_NUMREG_QUERY.Text
        End With
        
        Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
        Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione

Timer_ini = Timer
        
        Pcls_PrimaNota.PGestRegPN.ModificaRegistrazione
        
Timer_fin = Timer
Timer_diff_modificareg = Timer_diff_modificareg + (Timer_fin - Timer_ini)

        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.ModificaRegistrazione"
            Exit Sub
        End If
    Next Indice
    
TXT_TIMER_QUERY.AddItem Timer_diff_modificareg
    
    LST_AVANZAMENTO.AddItem String(40, "-")
    LST_AVANZAMENTO.ListIndex = LST_AVANZAMENTO.ListCount - 1
    
    DistruggiBOPrimaNota
    
Exit Sub
Err_CMD_QUERY_Click:
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
    
    '
    ' Valorizzazione campi
    '
    TXT_NUMEROMOVIMENTI.Text = 100
    TXT_NUM_COPPIE_D_A.Text = 8
    TXT_DATAREG.Text = "02/04/2010"
    TXT_CAUSALE.Text = "109"
    TXT_TOTALE.Text = 250
    
    TXT_CONTO1.Text = "1200040100"
    TXT_IMPORTO1.Text = 250
    CBO_SEGNO1.Text = 1
    
    TXT_CONTO2.Text = "1600010100"
    TXT_IMPORTO2.Text = 250
    CBO_SEGNO2.Text = 2
    
    '
    ' Variazione
    '
    TXT_NUMREG_QUERY.Text = "201000438090"
    TXT_NUMITER_QUERY.Text = 1
    
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
    
    TXT_TIMER_REG.Clear
    TXT_TIMER_RIGA.Clear
    TXT_TIMER_REGMOD.Clear
    
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
