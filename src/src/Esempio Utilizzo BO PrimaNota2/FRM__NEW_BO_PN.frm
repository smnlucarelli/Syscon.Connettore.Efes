VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_COMBOBOX.ocx"
Object = "{D8EB97B9-26FF-11D2-A1C0-0060082875F9}#6.16#0"; "TMS_EDITDEFCONTO.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.9#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.8#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.7#0"; "TMS_EDITDATE.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FRM__NEW_BO_PN 
   Caption         =   "Form1"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   10305
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   18177
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Inserimento mov. cont."
      TabPicture(0)   =   "FRM__NEW_BO_PN.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CMD_VIEW"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "CMD_INSERT"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(6)=   "TXT_NUMREG"
      Tab(0).Control(7)=   "Label5"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Inserimento fatture ( INTRA )"
      TabPicture(1)   =   "FRM__NEW_BO_PN.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TXT_NUMREG_INTRA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label58"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TXT_PROG"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CMD_INSERT_INTRA"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame10"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame11"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame8"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Variazione"
      TabPicture(2)   =   "FRM__NEW_BO_PN.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "CMD_QUERY"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "XML"
      TabPicture(3)   =   "FRM__NEW_BO_PN.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command1"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame8 
         Caption         =   "Avanzamento:"
         Height          =   1755
         Left            =   90
         TabIndex        =   132
         Top             =   8430
         Width           =   10095
         Begin VB.ListBox LST_AVANZ_IVA 
            Appearance      =   0  'Flat
            Height          =   1395
            Left            =   90
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   240
            Width           =   9945
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   675
         Left            =   -74760
         TabIndex        =   128
         Top             =   570
         Width           =   1785
      End
      Begin VB.Frame Frame11 
         Height          =   3885
         Left            =   60
         TabIndex        =   115
         Top             =   330
         Width           =   10125
         Begin VB.CheckBox CHK_DACONS_IVA 
            Caption         =   "Da consolidare"
            Height          =   285
            Left            =   150
            TabIndex        =   130
            Top             =   3450
            Width           =   2175
         End
         Begin VB.ListBox TXT_TIMER_REG_INTRA 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   3350
            Left            =   4320
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   360
            Width           =   1785
         End
         Begin VB.TextBox TXT_NUMREG_RM_INTRA 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2640
            TabIndex        =   120
            Text            =   "100"
            Top             =   1050
            Width           =   705
         End
         Begin VB.OptionButton CHK_REGMOD_OGNITOT_INTRA 
            Caption         =   "Invoca 'RegistraModifiche' ogni                  registrazioni"
            Height          =   255
            Left            =   90
            TabIndex        =   122
            Top             =   1080
            Width           =   4545
         End
         Begin VB.OptionButton CHK_REGMOD_UNICO_INTRA 
            Caption         =   "Invoca 'RegistraModifiche' una sola volta"
            Height          =   255
            Left            =   90
            TabIndex        =   121
            Top             =   780
            Value           =   -1  'True
            Width           =   3315
         End
         Begin VB.ListBox TXT_TIMER_RIGA_INTRA 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   3350
            Left            =   6150
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   360
            Width           =   1785
         End
         Begin VB.ListBox TXT_TIMER_REGMOD_INTRA 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   3350
            Left            =   7980
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   360
            Width           =   1785
         End
         Begin VB.CommandButton CMD_CLEAR_INTRA 
            Caption         =   "Cancella"
            Height          =   255
            Left            =   9300
            TabIndex        =   116
            Top             =   120
            Width           =   795
         End
         Begin PRJFW_EDIT.TxtEdit TXT_NUMEROMOVIMENTI_INTRA 
            Height          =   300
            Left            =   2520
            TabIndex        =   123
            Top             =   360
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin VB.Label Label29 
            Caption         =   "ins. riga"
            Height          =   225
            Left            =   6150
            TabIndex        =   127
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label28 
            Caption         =   "ins. registrazione"
            Height          =   225
            Left            =   4320
            TabIndex        =   126
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label Label27 
            Caption         =   "reg. mod."
            Height          =   225
            Left            =   7980
            TabIndex        =   125
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label26 
            Caption         =   "Numero di registrazioni da inserire"
            Height          =   225
            Left            =   90
            TabIndex        =   124
            Top             =   420
            Width           =   2445
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Dati relativi alla testata:"
         Height          =   1305
         Left            =   90
         TabIndex        =   94
         Top             =   4260
         Width           =   10095
         Begin PRJFW_EDITM.TXT_EDITM TXT_CAUSALE_INTRA 
            Height          =   300
            Left            =   4560
            TabIndex        =   95
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
            TabIndex        =   98
            Top             =   630
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            IsCalendario    =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREG_INTRA 
            Height          =   300
            Left            =   7680
            TabIndex        =   97
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
            TabIndex        =   101
            Top             =   630
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin VB.Label Label57 
            Caption         =   "Totale fattura"
            Height          =   225
            Left            =   3390
            TabIndex        =   112
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label56 
            Caption         =   "Num. doc. iniziale"
            Height          =   225
            Left            =   90
            TabIndex        =   111
            Top             =   660
            Width           =   1455
         End
         Begin VB.Label Label55 
            Caption         =   "Causale"
            Height          =   225
            Left            =   3390
            TabIndex        =   110
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label54 
            Caption         =   "Data reg."
            Height          =   225
            Left            =   6930
            TabIndex        =   109
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label53 
            Caption         =   "Codice ditta"
            Height          =   225
            Left            =   90
            TabIndex        =   108
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label52 
            Caption         =   "Cliente/forn."
            Height          =   225
            Left            =   90
            TabIndex        =   107
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label51 
            Caption         =   "Num. doc. orig."
            Height          =   225
            Left            =   3390
            TabIndex        =   106
            Top             =   630
            Width           =   1215
         End
         Begin VB.Label Label50 
            Caption         =   "Data doc."
            Height          =   225
            Left            =   6930
            TabIndex        =   105
            Top             =   660
            Width           =   975
         End
         Begin PRJFW_EDIT.TxtEdit TXT_DITTA_INTRA 
            Height          =   300
            Left            =   1590
            TabIndex        =   104
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
         Begin PRJFW_EDIT.TxtEdit TXT_NUMDOC 
            Height          =   300
            Left            =   1590
            TabIndex        =   103
            Top             =   630
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_TOTALE_INTRA 
            Height          =   300
            Left            =   4560
            TabIndex        =   102
            Top             =   960
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   529
            IsDbField       =   0   'False
            MaxChar         =   13
            CanRequired     =   0   'False
         End
         Begin PRJFW_EDIT.TxtEdit TXT_CLIFOR 
            Height          =   300
            Left            =   1590
            TabIndex        =   100
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
         Begin VB.Label Label49 
            Caption         =   "Segno"
            Height          =   225
            Left            =   6930
            TabIndex        =   99
            Top             =   990
            Width           =   675
         End
         Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNOCLIFOR 
            Height          =   315
            Left            =   7680
            TabIndex        =   96
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
      End
      Begin VB.Frame Frame9 
         Caption         =   "Dati relativi al dettaglio:"
         Height          =   1515
         Left            =   90
         TabIndex        =   76
         Top             =   6210
         Width           =   10095
         Begin PRJFW_EDITM.TXT_EDITM TXT_ALIQUOTA1 
            Height          =   300
            Left            =   2160
            TabIndex        =   77
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
         Begin PRJFW_EDITDEFCONTO.TxtEditDefConto TXT_CONTO1_INTRA 
            Height          =   300
            Left            =   2160
            TabIndex        =   78
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
            TabIndex        =   79
            Top             =   1080
            Width           =   1845
            _ExtentX        =   3281
            _ExtentY        =   529
            IsLookup        =   -1  'True
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPONIBILE1 
            Height          =   300
            Left            =   4860
            TabIndex        =   86
            Top             =   210
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            IsDbField       =   0   'False
            MaxWidth        =   11
            MaxChar         =   13
            CanRequired     =   0   'False
         End
         Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNOIVA 
            Height          =   315
            Left            =   7620
            TabIndex        =   80
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
         Begin VB.Label Label47 
            Caption         =   "Segno"
            Height          =   225
            Left            =   6990
            TabIndex        =   93
            Top             =   1140
            Width           =   675
         End
         Begin VB.Label Label46 
            Caption         =   "Conto di contropartita"
            Height          =   225
            Left            =   510
            TabIndex        =   92
            Top             =   300
            Width           =   1695
         End
         Begin VB.Label Label45 
            Caption         =   "Imponibile"
            Height          =   225
            Left            =   4080
            TabIndex        =   91
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label44 
            Caption         =   "Conto di contropartita IVA"
            Height          =   225
            Left            =   120
            TabIndex        =   90
            Top             =   1140
            Width           =   2025
         End
         Begin VB.Label Label43 
            Caption         =   "Importo"
            Height          =   225
            Left            =   4080
            TabIndex        =   89
            Top             =   1110
            Width           =   675
         End
         Begin VB.Label Label42 
            Caption         =   "Cod. aliquota"
            Height          =   225
            Left            =   510
            TabIndex        =   88
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label Label41 
            Caption         =   "Segno"
            Height          =   225
            Left            =   6990
            TabIndex        =   87
            Top             =   270
            Width           =   585
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOIVA 
            Height          =   300
            Left            =   4860
            TabIndex        =   85
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
            TabIndex        =   84
            Top             =   210
            Width           =   315
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   90
            X2              =   10000
            Y1              =   960
            Y2              =   960
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPOSTA1 
            Height          =   300
            Left            =   4860
            TabIndex        =   83
            Top             =   570
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   529
            IsDbField       =   0   'False
            MaxWidth        =   9
            MaxChar         =   13
            CanRequired     =   0   'False
         End
         Begin VB.Label Label34 
            Caption         =   "Imposta"
            Height          =   225
            Left            =   4080
            TabIndex        =   82
            Top             =   630
            Width           =   735
         End
         Begin PRJFW_COMBOBOX.TMS_COMBO CBO_SEGNO1_INTRA 
            Height          =   315
            Left            =   7620
            TabIndex        =   81
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
      End
      Begin VB.CommandButton CMD_INSERT_INTRA 
         Caption         =   "Esegui"
         Height          =   525
         Left            =   8850
         Picture         =   "FRM__NEW_BO_PN.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   7800
         Width           =   1305
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dati per autofattura"
         Height          =   555
         Left            =   90
         TabIndex        =   70
         Top             =   5610
         Width           =   10095
         Begin PRJFW_EDITM.TXT_EDITM TXT_SEZAUTO 
            Height          =   300
            Left            =   1230
            TabIndex        =   71
            Top             =   210
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   529
            IsLookup        =   0   'False
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   4
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            NumRighe        =   0
            MaxWidth        =   4
            CanRequired     =   0   'False
         End
         Begin VB.Label Label31 
            Caption         =   "Num. doc."
            Height          =   225
            Left            =   2670
            TabIndex        =   74
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label32 
            Caption         =   "Sezionale"
            Height          =   225
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   825
         End
         Begin PRJFW_EDIT.TxtEdit TXT_NUMDOCAUTO 
            Height          =   300
            Left            =   3540
            TabIndex        =   72
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1605
         Left            =   -74940
         TabIndex        =   62
         Top             =   360
         Width           =   10125
         Begin VB.CommandButton CMD_CLEAR_QUERY 
            Caption         =   "Cancella"
            Height          =   255
            Left            =   9300
            TabIndex        =   64
            Top             =   120
            Width           =   795
         End
         Begin VB.ListBox TXT_TIMER_QUERY 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   1200
            Left            =   7650
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   360
            Width           =   1785
         End
         Begin VB.Label Label23 
            Caption         =   "Numero di iterazioni"
            Height          =   225
            Left            =   90
            TabIndex        =   69
            Top             =   660
            Width           =   1605
         End
         Begin VB.Label Label25 
            Caption         =   "modifica registrazione"
            Height          =   225
            Left            =   7650
            TabIndex        =   68
            Top             =   120
            Width           =   1575
         End
         Begin PRJFW_EDIT.TxtEdit TXT_NUMREG_QUERY 
            Height          =   300
            Left            =   1980
            TabIndex        =   67
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
         Begin PRJFW_EDIT.TxtEdit TXT_NUMITER_QUERY 
            Height          =   300
            Left            =   1980
            TabIndex        =   66
            Top             =   630
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin VB.Label Label24 
            Caption         =   "NUMREG da interrogare"
            Height          =   225
            Left            =   90
            TabIndex        =   65
            Top             =   270
            Width           =   1815
         End
      End
      Begin VB.CommandButton CMD_QUERY 
         Caption         =   "Esegui"
         Height          =   525
         Left            =   -66120
         Picture         =   "FRM__NEW_BO_PN.frx":01BA
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   2010
         Width           =   1305
      End
      Begin VB.CommandButton CMD_VIEW 
         Caption         =   "Visualizza"
         Height          =   525
         Left            =   -66120
         Picture         =   "FRM__NEW_BO_PN.frx":0304
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   8610
         Width           =   1305
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dati relativi alla testata:"
         Height          =   2025
         Left            =   -74940
         TabIndex        =   39
         Top             =   1980
         Width           =   10125
         Begin VB.CheckBox CHK_DACONS 
            Caption         =   "Da consolidare"
            Height          =   285
            Left            =   150
            TabIndex        =   129
            Top             =   330
            Width           =   2175
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_CAUSALE 
            Height          =   300
            Left            =   1350
            TabIndex        =   40
            Top             =   1080
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
            TabIndex        =   41
            Top             =   1080
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
            Top             =   900
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
            Top             =   660
            Width           =   1905
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATARATEOCORR 
            Height          =   300
            Left            =   8130
            TabIndex        =   51
            Top             =   1500
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
            Top             =   1260
            Width           =   1935
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREG 
            Height          =   300
            Left            =   4410
            TabIndex        =   49
            Top             =   750
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
            Top             =   1470
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
            Top             =   720
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
            Top             =   1530
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Totale registrazione"
            Height          =   225
            Left            =   2880
            TabIndex        =   45
            Top             =   1140
            Width           =   1605
         End
         Begin VB.Label Label3 
            Caption         =   "Causale"
            Height          =   225
            Left            =   90
            TabIndex        =   44
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Data reg."
            Height          =   225
            Left            =   2880
            TabIndex        =   43
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Codice ditta"
            Height          =   225
            Left            =   90
            TabIndex        =   42
            Top             =   810
            Width           =   975
         End
      End
      Begin VB.CommandButton CMD_INSERT 
         Caption         =   "Esegui"
         Height          =   525
         Left            =   -66120
         Picture         =   "FRM__NEW_BO_PN.frx":044E
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   6660
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dati relativi al dettaglio:"
         Height          =   2565
         Left            =   -74940
         TabIndex        =   14
         Top             =   4050
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
         Left            =   -74940
         TabIndex        =   12
         Top             =   7260
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
         Left            =   -74940
         TabIndex        =   1
         Top             =   360
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
      Begin PRJFW_EDIT.TxtEdit TXT_PROG 
         Height          =   300
         Left            =   3030
         TabIndex        =   131
         Top             =   7800
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
      Begin VB.Label Label58 
         Caption         =   "Ultimo numero registrazione assegnato"
         Height          =   225
         Left            =   120
         TabIndex        =   114
         Top             =   7830
         Width           =   2745
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG_INTRA 
         Height          =   300
         Left            =   120
         TabIndex        =   113
         Top             =   8070
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
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
         Height          =   300
         Left            =   -74940
         TabIndex        =   55
         Top             =   6930
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
         Left            =   -74910
         TabIndex        =   54
         Top             =   6690
         Width           =   2745
      End
   End
End
Attribute VB_Name = "FRM__NEW_BO_PN"
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

Private Sub CMD_CLEAR_INTRA_Click()
    On Error Resume Next
    
    TXT_TIMER_REG_INTRA.Clear
    TXT_TIMER_RIGA_INTRA.Clear
    TXT_TIMER_REGMOD_INTRA.Clear
    
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
        
        Pcls_PrimaNota.PGestRegPN.CPInput.InserimentoMultiplo = True
        
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
            
            If CHK_DACONS.Value = vbChecked Then
                .IndicatoreTipoMovimento = DaConsolidare
            Else
                .IndicatoreTipoMovimento = Consolidato
            End If
            
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
'            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_GestRegPN.ModificaRegistrazione"
'            Exit Sub
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
    TXT_DITTA_INTRA.Text = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    '
    ' Setto le proprietà dei conti
    '
    SettaProprietaConti
    
    '
    ' Valorizzazione campi - inserimento mov. cont.
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
    
    '
    ' inserimento fatt. acq. intra
    '
    TXT_NUMEROMOVIMENTI_INTRA.Text = 1
    TXT_NUMREG_RM_INTRA.Text = 100
    
    TXT_NUMDOC.Text = "123"
    TXT_CLIFOR.Text = 1 ' 3 = fornitore intra
    TXT_CAUSALE_INTRA.Text = "1"
    TXT_CAUSALE_INTRA_Validate False
    TXT_NUMDOCORIG.Text = "orig"
    TXT_TOTALE_INTRA.Text = 1200
    TXT_DATAREG_INTRA.Text = "13/10/2012"
    TXT_DATADOC.Text = "12/10/2012"
    
    TXT_SEZAUTO.Text = "00"
    TXT_NUMDOCAUTO.Text = 0 ' no autofattura
    
    TXT_CONTO1_INTRA.Text = "1200040100"
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
    ' Carico il combo del segno
    '
    CBO_SEGNO1.AddItemData "Dare", 1
    CBO_SEGNO1.AddItemData "Avere", 2
    
    CBO_SEGNO2.AddItemData "Dare", 1
    CBO_SEGNO2.AddItemData "Avere", 2
    
    TXT_TIMER_REG.Clear
    TXT_TIMER_RIGA.Clear
    TXT_TIMER_REGMOD.Clear
    
    '
    ' Fattura acquisto INTRA
    '
    CBO_SEGNOCLIFOR.AddItemData "Dare", 1
    CBO_SEGNOCLIFOR.AddItemData "Avere", 2
    
    CBO_SEGNO1_INTRA.AddItemData "Dare", 1
    CBO_SEGNO1_INTRA.AddItemData "Avere", 2
    
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
    
    Set TXT_CONTO1_INTRA.Connessione = Connessione
    TXT_CONTO1_INTRA.CodicePdC = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPdc
    TXT_CONTO1_INTRA.Ditta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
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

Private Sub TXT_CAUSALE_INTRA_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
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

Private Sub TXT_CAUSALE_INTRA_Validate(Cancel As Boolean)
    Dim TipoRegistro        As TipoRegistroEnum
    
    On Error Resume Next
    
    TipoRegistro = GetTipoRegistro(TXT_CAUSALE_INTRA.Text)
    Select Case TipoRegistro
        Case TipoRegistroEnum.RegistroVendite
            CBO_SEGNOCLIFOR.Text = 1 'dare
            CBO_SEGNO1_INTRA.Text = 2 'avere
            CBO_SEGNOIVA.Text = 2 'avere
            
            TXT_CONTOIVA.Text = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoIvaVendite
        Case TipoRegistroEnum.RegistroAcquisti
            CBO_SEGNOCLIFOR.Text = 2 'avere
            CBO_SEGNO1_INTRA.Text = 1 'dare
            CBO_SEGNOIVA.Text = 1 'dare
            
            TXT_CONTOIVA.Text = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoIvaAcquisti
    End Select
    Err.Clear
End Sub

Private Sub TXT_CONTO1_INTRA_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CONTO1_INTRA_StartLookup
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CONTO1_INTRA_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CONTO1_INTRA_StartLookup"
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

Private Sub CMD_INSERT_INTRA_Click()
    Dim Indice              As Variant
    Dim MastroClienti       As Variant
    Dim MastroFornitori     As Variant
    Dim TipoRegistro        As TipoRegistroEnum
    Dim RecSetScadenze      As ADODB.Recordset
    Dim StatoFinale         As StatoFinaleEnum
    Dim NumProt             As Variant
    
    On Error GoTo Err_CMD_INSERT_INTRA_Click
    
    '
    ' Pulisco il list box che segnala le operazioni
    '
    LST_AVANZ_IVA.Clear
    LST_AVANZ_IVA.Refresh
    
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
Timer_diff_registramodifiche = 0
    
    For Indice = 1 To NVL(TXT_NUMEROMOVIMENTI_INTRA.Text, 0)
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
            
            .CodiceDitta = TXT_DITTA_INTRA.Text
            .DataRegistrazione = TXT_DATAREG_INTRA.Text
            .NumeroRegistrazione = ""
            .CodiceCausale = TXT_CAUSALE_INTRA.Text
            .NumeroDocumento = NumProt '  TXT_NUMDOC.Text
            .NumeroPartita = NumProt ' TXT_NUMDOC.Text
            .DataRegIva = TXT_DATAREG_INTRA.Text
            '
            ' In questo caso vengono trattate solo le causali 1 (fattura di vendita)
            ' e 31 (fattura di acquisto)
            '
            TipoRegistro = GetTipoRegistro(TXT_CAUSALE_INTRA.Text)
            Select Case TipoRegistro
                Case RegistroAcquisti
                    .ContoCliFor = Left(MastroFornitori, 2) & Fill0(TXT_CLIFOR.Text, 8)
                Case RegistroVendite
                    .ContoCliFor = Left(MastroClienti, 2) & Fill0(TXT_CLIFOR.Text, 8)
            End Select
            
            .NumeroDocumentoOrigine = TXT_NUMDOCORIG.Text
            .DataDocumentoOrigine = TXT_DATADOC.Text
            .ImportoDocumento = TXT_TOTALE_INTRA.Text
            
            If CHK_DACONS_IVA.Value = vbChecked Then
                .IndicatoreTipoMovimento = DaConsolidare
            Else
                .IndicatoreTipoMovimento = Consolidato
            End If
            
            .TipoDocumento = Pcls_PrimaNota.PGestRegPN.GetTipoDocumento(.CodiceCausale)
            
            '
            ' Dati per autofattura
            '
            .CodiceSezionaleAutofattura = NVL(TXT_SEZAUTO.Text, "00")
            .NumeroDocumentoAutofattura = NVL(TXT_NUMDOCAUTO.Text, 0)
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
        
        TXT_NUMREG_INTRA.Text = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
        
        TXT_PROG.Text = Indice
        
        '
        ' Inserimento riga del cli/for (partita)
        '
        With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
            .CodiceDitta = TXT_DITTA_INTRA.Text
            .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
            .Conto = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.ContoCliFor
            .Importo = TXT_TOTALE_INTRA.Text
            .Segno = CBO_SEGNOCLIFOR.Text
            .IndicatoreTipoOperazione = DocumentoIVA_MovimentoPartitaDocumento
        End With
        
        Pcls_PrimaNota.PGestRegPN.InserisciRiga
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
            Exit Sub
        End If
        
        '
        ' Inserimento prima riga di contropartita
        '
        With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
            .CodiceDitta = TXT_DITTA_INTRA.Text
            .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
            .Conto = TXT_CONTO1_INTRA.Text
            .Importo = TXT_IMPONIBILE1.Text
            .Imponibile = TXT_IMPONIBILE1.Text
            .Segno = CBO_SEGNO1_INTRA.Text
            .CodiceAliquota = TXT_ALIQUOTA1.Text
            .Imposta = TXT_IMPOSTA1.Text
            .ImpostaND = 0
            .IndicatoreTipoOperazione = DocumentoIVA_ContropartiteSuFattura
            .CausaleIva = Null
        End With
        
        Pcls_PrimaNota.PGestRegPN.InserisciRiga
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
            Exit Sub
        End If
        
        '
        ' Inserimento riga di contropartita IVA
        '
        If NVL(TXT_IMPORTOIVA.Text, 0) <> 0 Then
            With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
                .CodiceDitta = TXT_DITTA_INTRA.Text
                .NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
                .Conto = TXT_CONTOIVA.Text
                .Importo = TXT_IMPORTOIVA.Text
                .Segno = CBO_SEGNOIVA.Text
                .CodiceAliquota = Null
                .Imposta = 0
                .ImpostaND = 0
                .IndicatoreTipoOperazione = DocumentoIVA_IvaSuFatturaCorrispettivo
            End With
            
            Pcls_PrimaNota.PGestRegPN.InserisciRiga
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
            LST_AVANZ_IVA.AddItem "ERRORE NON BLOCCANTE: " & Pcls_PrimaNota.PGestRegPN.StatoNonBloccante & " - " & Pcls_PrimaNota.PGestRegPN.ErroreNonBloccante
            LST_AVANZ_IVA.ListIndex = LST_AVANZ_IVA.ListCount - 1
        End If
        
        Set RecSetScadenze = Nothing
    Next
    
TXT_TIMER_REG_INTRA.AddItem Timer_diff_ins_reg
'TXT_TIMER_RIGA_INTRA.AddItem Timer_diff_ins_riga
TXT_TIMER_REGMOD_INTRA.AddItem Timer_diff_registramodifiche
    
    MsgBox "Le registrazioni sono state inserite"
    
Exit Sub
Err_CMD_INSERT_INTRA_Click:
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
    Pobj_PNota.IndicatoreTipoRegistro = GetTipoRegistro(TXT_CAUSALE_INTRA.Text)
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
    
    PClsScadenze.Conto = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroClienti ' Fornitori ' Clienti
    PClsScadenze.TipoCliFor = NVL(Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_TIPOCF_CG44").Value, 0)
    PClsScadenze.CodiceCliFor = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CLIFOR_CG44").Value
    PClsScadenze.CondPagamento = GetCondizionePagamento(PClsScadenze.TipoCliFor, PClsScadenze.CodiceCliFor)
    
    PClsScadenze.DareAvere = 1 ' Cliente -> Dare (= 1)
    
    PClsScadenze.TipoElaborazioneScadenze = Inizializzazione
    
    PClsScadenze.Ditta = TXT_DITTA_INTRA.Text
    PClsScadenze.NumeroRigaContabile = 1 ' riferimento al numero riga contabile che apre la partita
    
    PClsScadenze.NumeroRegistrazione = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMREG").Value
    PClsScadenze.CausaleContabile = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_CODICE_CG33").Value
    PClsScadenze.NumeroPartita = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMDOC").Value
    PClsScadenze.SezionalePartita = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_SEZIONALE").Value
    PClsScadenze.PartitaBis = 0
    PClsScadenze.DataRegistrazione = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DATAREG").Value
    PClsScadenze.NumeroDocumento = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMDOC").Value
    PClsScadenze.Sezionale = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_SEZIONALE").Value
    PClsScadenze.DocumentoBis = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_FLGDOCBIS").Value
    PClsScadenze.NumeroDocumentoOrigine = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_NUMDOCORIG").Value
    PClsScadenze.FlagAcconto = 0
    PClsScadenze.DescrCausaleContabile = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_DESCAUSALE").Value
    PClsScadenze.IndTipoMov = Pcls_PrimaNota.PGestRegPN.CPOutput.RecSetTestata.Fields("CG41_INDTIPOMOV").Value
    
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






'-------------------------------------
'                XML
'-------------------------------------
Private Sub Command1_Click()
    BulkLoadFromStream
End Sub

Private Sub BulkLoadFromStream()
'''    Dim objBL           As SQLXMLBULKLOADLib.SQLXMLBulkLoad4
'''    Dim objCmd          As ADODB.Command
'''    Dim objStrmOut      As ADODB.Stream
'''
''''    Const adCRLF = -1
''''    Const adExecuteStream = 1024
'''
'''    Set objBL = New SQLXMLBULKLOADLib.SQLXMLBulkLoad4
'''    Set objCmd = New ADODB.Command
'''    Set objStrmOut = New ADODB.Stream
'''
'''    objBL.ConnectionString = StrConnect
'''    objBL.ErrorLogFile = "c:\XMLerror.log"
'''    objBL.CheckConstraints = True
'''    objBL.SchemaGen = True
'''    objBL.SGDropTables = True
'''    objBL.XMLFragment = True
'''    Set objCmd.ActiveConnection = Connessione
'''    objCmd.CommandText = "SELECT * FROM CG41_PRIMANOTA_NAME FOR XML AUTO, ELEMENTS"
'''
'''    ' Open the return stream and execute the command.
'''    objStrmOut.Open
'''    objStrmOut.LineSeparator = adCRLF
'''    objCmd.Properties("Output Stream").Value = objStrmOut
'''    objCmd.Execute , , 1024  ' adExecuteStream
'''    objStrmOut.Position = 0
'''
'''
'''    objStrmOut.SaveToFile "C:\TestXml.xml", adSaveCreateOverWrite
'''
'''    ' Execute bulk load. Read source XML data from the stream.
'''    '' objBL.Execute "SampleSchema.xml", objStrmOut
'''
'''    Set objBL = Nothing
End Sub
