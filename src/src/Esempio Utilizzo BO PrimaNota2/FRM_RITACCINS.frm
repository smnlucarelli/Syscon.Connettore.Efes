VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_COMBOBOX.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.10#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.9#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.8#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.7#0"; "TMS_EDITDATE.ocx"
Object = "{F53BE214-7AC6-11D0-9B0E-006097A80EFD}#6.7#0"; "TMS_LABEL.ocx"
Object = "{0EF4EAD5-2617-11D2-A1C0-0060082875F9}#5.9#0"; "TMS_CHECKBOX.ocx"
Object = "{F2DC983F-61F7-11D2-AE21-00A0244C5B50}#3.9#0"; "TMS_EDITDATEM.ocx"
Begin VB.Form FRM_RITACCINS 
   Caption         =   "Gestione ritenute acconto - inserimento"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   Icon            =   "FRM_RITACCINS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Dati relativi al pagamento:"
      Height          =   2805
      Left            =   30
      TabIndex        =   0
      Top             =   4140
      Width           =   11715
      Begin VB.CommandButton BUT_INSERTPAG 
         Caption         =   "Inserisci pagamento"
         Height          =   345
         Left            =   7980
         Picture         =   "FRM_RITACCINS.frx":27A2
         TabIndex        =   55
         Top             =   2400
         Width           =   1785
      End
      Begin VB.CommandButton BUT_REGISTRAMODIFICHEPAG 
         Caption         =   "Registra Modifiche"
         Height          =   345
         Left            =   9840
         Picture         =   "FRM_RITACCINS.frx":28EC
         TabIndex        =   56
         Top             =   2400
         Width           =   1785
      End
      Begin VB.CommandButton BUT_PROPONIIMPORTI 
         Caption         =   "Proponi importi"
         Height          =   615
         Left            =   8970
         Picture         =   "FRM_RITACCINS.frx":2A36
         TabIndex        =   40
         Top             =   210
         Width           =   855
      End
      Begin VB.CommandButton BUT_PROPONIPAGATO 
         Caption         =   "Proponi pagato"
         Height          =   615
         Left            =   3060
         Picture         =   "FRM_RITACCINS.frx":2B80
         TabIndex        =   35
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton RDB_TOT 
         Caption         =   "Ritenute totali"
         Height          =   195
         Left            =   6780
         TabIndex        =   39
         Top             =   600
         Width           =   2025
      End
      Begin VB.OptionButton RDB_PROP 
         Caption         =   "Ritenute in proporzione"
         Height          =   195
         Left            =   6780
         TabIndex        =   38
         Top             =   270
         Value           =   -1  'True
         Width           =   2025
      End
      Begin PRJFW_EDITDATEM.EditDateM TXT_DATAPAG 
         Height          =   300
         Left            =   1410
         TabIndex        =   41
         Top             =   1020
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         IsCalendario    =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         Caption         =   "Data pagamento"
         Object.Tag             =   "Data pagamento fattura con ritenute"
         Formato         =   "Medium Date"
      End
      Begin PRJFW_CHECKBOX.TMS_CHECKBOX CHK_STANDALONEPAG 
         Height          =   300
         Left            =   120
         TabIndex        =   54
         Top             =   2400
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Movimento stand-alone"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREGPAG 
         Height          =   300
         Left            =   6180
         TabIndex        =   102
         Top             =   2430
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
      Begin VB.Label Label17 
         Caption         =   "Num. reg. pagamento"
         Height          =   225
         Left            =   4575
         TabIndex        =   101
         Top             =   2460
         Width           =   1635
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPON_RP 
         Height          =   300
         Left            =   9630
         TabIndex        =   51
         Top             =   1350
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Imponibile RP"
         Object.Tag             =   "Imponibile soggetto a ritenuta previdenziale"
         MaxWidth        =   11
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_FRANCHIGIA 
         Height          =   300
         Left            =   9630
         TabIndex        =   50
         Top             =   1020
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Franchigia"
         Object.Tag             =   "Importo della franchigia sull'imponibile ritenuta previdenziale"
         MaxWidth        =   11
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_CONTCDITTA 
         Height          =   300
         Left            =   9630
         TabIndex        =   53
         Top             =   2010
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Importo c/ditta"
         Object.Tag             =   "Importo ritenuta previdenziale carico ditta"
         MaxWidth        =   11
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_CONTCPERC 
         Height          =   300
         Left            =   9630
         TabIndex        =   52
         Top             =   1680
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Importo c/percipiente"
         Object.Tag             =   "Importo ritenuta previdenziale carico percipiente"
         MaxWidth        =   11
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL37 
         Height          =   300
         Left            =   7710
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         Caption         =   "Imponibile soggetto R.P."
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL38 
         Height          =   300
         Left            =   7710
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         Caption         =   "Franchigia R.P."
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL40 
         Height          =   300
         Left            =   7710
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   529
         Caption         =   "Importo R.P. c/percipiente"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL3 
         Height          =   300
         Left            =   7710
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         Caption         =   "Importo R.P. c/ditta"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO_RA 
         Height          =   300
         Left            =   5400
         TabIndex        =   47
         Top             =   1350
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Importo RA"
         Object.Tag             =   "Importo ritenuta d'acconto"
         MaxWidth        =   11
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPON_RA 
         Height          =   300
         Left            =   5400
         TabIndex        =   46
         Top             =   1020
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Imponibile soggetto RA"
         Object.Tag             =   "Imponibile soggetto R.A."
         MaxWidth        =   11
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_RIMBORSI 
         Height          =   300
         Left            =   5400
         TabIndex        =   49
         Top             =   2010
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Importo non soggetto"
         Object.Tag             =   "Importo non soggetto a ritenuta d'acconto"
         MaxWidth        =   11
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOALRIT 
         Height          =   300
         Left            =   5400
         TabIndex        =   48
         Top             =   1680
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Importo altre ritenute"
         Object.Tag             =   "Importo altre ritenute"
         MaxWidth        =   11
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL7 
         Height          =   300
         Left            =   3480
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         Caption         =   "Importo altre ritenute"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL34 
         Height          =   195
         Left            =   3480
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         Caption         =   "Importo R.A."
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL32 
         Height          =   285
         Left            =   3480
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         Caption         =   "Imponibile soggetto R.A."
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL42 
         Height          =   195
         Left            =   3480
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   529
         Caption         =   "Importo non soggetto"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL4 
         Height          =   300
         Left            =   150
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         Caption         =   "Provv. non sogg."
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL2 
         Height          =   300
         Left            =   150
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         Caption         =   "Contr. integr."
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_CONTRIBINT 
         Height          =   300
         Left            =   1410
         TabIndex        =   44
         Top             =   1680
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Contributo integrativo"
         Object.Tag             =   "Importo contributo integrativo"
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PROVVNONSOG 
         Height          =   300
         Left            =   1410
         TabIndex        =   45
         Top             =   2010
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Provvigioni non soggette"
         Object.Tag             =   "Provvigioni non soggette a contributo"
         MaxChar         =   13
         FormatMask      =   "##,###,###,##0.00"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL1 
         Height          =   255
         Left            =   150
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "Data pagamento"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL29 
         Height          =   300
         Left            =   150
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         Caption         =   "Mese/Anno comp."
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_MESECOMP 
         Height          =   300
         Left            =   1410
         TabIndex        =   42
         Top             =   1350
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Mese competenza"
         Object.Tag             =   "Mese competenza pagamento fattura con ritenute"
         MaxWidth        =   2
         MaxChar         =   2
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_ANNOCOMP 
         Height          =   300
         Left            =   1890
         TabIndex        =   43
         Top             =   1350
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Anno competenza"
         Object.Tag             =   "Anno competenza pagamento fattura con ritenuta"
         MaxWidth        =   4
         MaxChar         =   4
         CanRequired     =   0   'False
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   150
         X2              =   11630
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   4020
         X2              =   4020
         Y1              =   180
         Y2              =   950
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NREGPERPAG 
         Height          =   300
         Left            =   1470
         TabIndex        =   34
         Top             =   570
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
      Begin VB.Label Label8 
         Caption         =   "Num. reg. ritenuta"
         Height          =   225
         Left            =   135
         TabIndex        =   87
         Top             =   600
         Width           =   1305
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTAPAG 
         Height          =   300
         Left            =   1470
         TabIndex        =   33
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         MaxChar         =   6
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   6
         CanRequired     =   0   'False
         Allineamento    =   1
      End
      Begin VB.Label Label5 
         Caption         =   "Ditta"
         Height          =   225
         Left            =   150
         TabIndex        =   86
         Top             =   300
         Width           =   975
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PAGATO 
         Height          =   300
         Left            =   4890
         TabIndex        =   36
         Top             =   210
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label9 
         Caption         =   "Pagato"
         Height          =   225
         Left            =   4140
         TabIndex        =   58
         Top             =   270
         Width           =   645
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_ABBUONO 
         Height          =   300
         Left            =   4890
         TabIndex        =   37
         Top             =   540
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label23 
         Caption         =   "Abbuono"
         Height          =   225
         Left            =   4140
         TabIndex        =   57
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dati relativi alla ritenuta:"
      Height          =   4035
      Left            =   30
      TabIndex        =   59
      Top             =   60
      Width           =   11715
      Begin VB.CommandButton BUT_REGISTRAMODIFICHE 
         Caption         =   "Registra Modifiche"
         Height          =   345
         Left            =   9660
         Picture         =   "FRM_RITACCINS.frx":2CCA
         TabIndex        =   32
         Top             =   2700
         Width           =   1785
      End
      Begin VB.CommandButton BUT_INSERT 
         Caption         =   "Inserisci mov. ritenuta"
         Height          =   345
         Left            =   9660
         Picture         =   "FRM_RITACCINS.frx":2E14
         TabIndex        =   31
         Top             =   2310
         Width           =   1785
      End
      Begin VB.CommandButton BUT_CALCOLAIMPORTI 
         Caption         =   "Calcola importi"
         Height          =   375
         Left            =   4950
         Picture         =   "FRM_RITACCINS.frx":2F5E
         TabIndex        =   16
         Top             =   1230
         Width           =   1275
      End
      Begin VB.CommandButton BUT_INFOCAUPREST 
         Caption         =   "Info caus. prestaz."
         Height          =   495
         Left            =   2490
         Picture         =   "FRM_RITACCINS.frx":30A8
         TabIndex        =   5
         Top             =   1050
         Width           =   855
      End
      Begin VB.CommandButton BUT_INFOANAGGEN 
         Caption         =   "Informazioni anagrafica"
         Height          =   525
         Left            =   1860
         Picture         =   "FRM_RITACCINS.frx":31F2
         TabIndex        =   3
         Top             =   360
         Width           =   1005
      End
      Begin PRJFW_EDITM.TXT_EDITM TXT_CAUSPREST 
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Top             =   1140
         Width           =   1140
         _ExtentX        =   1984
         _ExtentY        =   529
         IsLookup        =   -1  'True
         DisplayFormat   =   "Maiuscolo"
         MaxChar         =   4
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         NumRighe        =   0
         MaxWidth        =   5
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITM.TXT_EDITM TXT_CODTRIBUTO 
         Height          =   300
         Left            =   1320
         TabIndex        =   9
         Top             =   2595
         Width           =   1020
         _ExtentX        =   1773
         _ExtentY        =   529
         IsLookup        =   -1  'True
         DisplayFormat   =   "Maiuscolo"
         MaxChar         =   4
         Carattere       =   0   'False
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         NumRighe        =   0
         MaxWidth        =   4
         CanRequired     =   0   'False
         LookupColumn    =   1
         CanReturnRecordSet=   -1  'True
      End
      Begin PRJFW_EDITM.TXT_EDITM TXT_CODTRIBRP 
         Height          =   300
         Left            =   1320
         TabIndex        =   11
         Top             =   3285
         Width           =   1020
         _ExtentX        =   1773
         _ExtentY        =   529
         IsLookup        =   -1  'True
         DisplayFormat   =   "Maiuscolo"
         MaxChar         =   4
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         IsDecode        =   -1  'True
         NumRighe        =   0
         MaxWidth        =   4
         CanRequired     =   0   'False
         LookupColumn    =   1
         CanReturnRecordSet=   -1  'True
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PERCRIPAZ 
         Height          =   300
         Left            =   2640
         TabIndex        =   104
         Tag             =   "Percentuale Inps"
         Top             =   2940
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   4
         MaxChar         =   5
         FormatMask      =   "###.00;-###.00;;"
         CanRequired     =   0   'False
         GetTextMode     =   1
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL5 
         Height          =   300
         Left            =   2130
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   2970
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   529
         Caption         =   "% Az."
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   60
         X2              =   3300
         Y1              =   1020
         Y2              =   1020
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_ALTRERITENUTE 
         Height          =   300
         Left            =   7890
         TabIndex        =   26
         Top             =   3000
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label16 
         Caption         =   "Altre ritenute"
         Height          =   225
         Left            =   6435
         TabIndex        =   100
         Top             =   3060
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Num. reg. testata ritenuta"
         Height          =   225
         Left            =   9555
         TabIndex        =   85
         Top             =   1650
         Width           =   1965
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
         Height          =   300
         Left            =   9540
         TabIndex        =   84
         Top             =   1890
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
      Begin PRJFW_CHECKBOX.TMS_CHECKBOX CHK_BIS 
         Height          =   300
         Left            =   10290
         TabIndex        =   30
         Top             =   1260
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Bis"
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   2085
         Left            =   9480
         Top             =   180
         Width           =   2175
      End
      Begin PRJFW_CHECKBOX.TMS_CHECKBOX CHK_STANDALONE 
         Height          =   300
         Left            =   9510
         TabIndex        =   27
         Top             =   270
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   529
         IsDbField       =   0   'False
         Caption         =   "Movimento stand-alone"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOIVA 
         Height          =   300
         Left            =   4860
         TabIndex        =   15
         Top             =   870
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label14 
         Caption         =   "Importo IVA"
         Height          =   225
         Left            =   3420
         TabIndex        =   83
         Top             =   900
         Width           =   975
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOSOGG 
         Height          =   300
         Left            =   4860
         TabIndex        =   14
         Top             =   540
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label11 
         Caption         =   "Importo sogg."
         Height          =   225
         Left            =   3420
         TabIndex        =   82
         Top             =   570
         Width           =   1005
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOREG 
         Height          =   300
         Left            =   4860
         TabIndex        =   13
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label10 
         Caption         =   "Importo reg."
         Height          =   255
         Left            =   3420
         TabIndex        =   81
         Top             =   240
         Width           =   945
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   9420
         X2              =   9420
         Y1              =   180
         Y2              =   3900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   3360
         X2              =   3360
         Y1              =   180
         Y2              =   3900
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOCDITTA 
         Height          =   300
         Left            =   7890
         TabIndex        =   24
         Top             =   2340
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label37 
         Caption         =   "Importo carico ditta"
         Height          =   225
         Left            =   6435
         TabIndex        =   80
         Top             =   2385
         Width           =   1425
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTOCPERC 
         Height          =   300
         Left            =   7890
         TabIndex        =   23
         Top             =   2010
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label36 
         Caption         =   "Importo carico perc."
         Height          =   225
         Left            =   6435
         TabIndex        =   79
         Top             =   2055
         Width           =   1455
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPONIBILERP 
         Height          =   300
         Left            =   7890
         TabIndex        =   22
         Top             =   1680
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label35 
         Caption         =   "Imponibile sogg. RP"
         Height          =   225
         Left            =   6435
         TabIndex        =   78
         Top             =   1725
         Width           =   1425
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_FRANCHIGIARP 
         Height          =   300
         Left            =   4860
         TabIndex        =   21
         Top             =   3000
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label34 
         Caption         =   "Franchigia RP"
         Height          =   225
         Left            =   3405
         TabIndex        =   77
         Top             =   3045
         Width           =   1095
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTORA 
         Height          =   300
         Left            =   4860
         TabIndex        =   20
         Top             =   2670
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label33 
         Caption         =   "Importo RA"
         Height          =   225
         Left            =   3405
         TabIndex        =   76
         Top             =   2715
         Width           =   915
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_CONTRIBINTEGR 
         Height          =   300
         Left            =   4860
         TabIndex        =   19
         Top             =   2340
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label32 
         Caption         =   "Contrib. integrativo"
         Height          =   225
         Left            =   3405
         TabIndex        =   75
         Top             =   2370
         Width           =   1395
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPONIBILERA 
         Height          =   300
         Left            =   4860
         TabIndex        =   18
         Top             =   2010
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label31 
         Caption         =   "Imponibile sogg. RA"
         Height          =   225
         Left            =   3405
         TabIndex        =   74
         Top             =   2040
         Width           =   1455
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PROVVNONSOGG 
         Height          =   300
         Left            =   4860
         TabIndex        =   17
         Top             =   1680
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label7 
         Caption         =   "Provv. non sogg."
         Height          =   255
         Left            =   3405
         TabIndex        =   73
         Top             =   1710
         Width           =   1395
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PERCBASEIMP 
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   1920
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         Default         =   "100"
         IsDbField       =   0   'False
         MaxWidth        =   4
         MaxChar         =   5
         FormatMask      =   "###.00;-###.00;;"
         CanRequired     =   0   'False
      End
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_PERCBASEIMP 
         Height          =   300
         Left            =   120
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1950
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         Caption         =   "% Base impon."
      End
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_CODTRIBUTO 
         Height          =   300
         Left            =   120
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2625
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "Cod. tributo"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_PERCRA 
         Height          =   300
         Left            =   120
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2280
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "% Rit.acc."
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PERCRA 
         Height          =   300
         Left            =   1320
         TabIndex        =   8
         Tag             =   "Percentuale Inps"
         Top             =   2250
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   4
         MaxChar         =   5
         FormatMask      =   "###.00;-###.00;;"
         CanRequired     =   0   'False
         GetTextMode     =   1
      End
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_PERCTRIB 
         Height          =   300
         Left            =   120
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2970
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "% Rit.prev."
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PERCINPS 
         Height          =   300
         Left            =   1320
         TabIndex        =   10
         Tag             =   "Percentuale Inps"
         Top             =   2940
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   4
         MaxChar         =   5
         FormatMask      =   "###.00;-###.00;;"
         CanRequired     =   0   'False
         GetTextMode     =   1
      End
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_CODTRIBRP 
         Height          =   300
         Left            =   120
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   3315
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "Cod. tributo"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PERCCI 
         Height          =   300
         Left            =   1320
         TabIndex        =   12
         Tag             =   "Percentuale Inps"
         Top             =   3615
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         DBField         =   "CG15_PERCCI"
         Caption         =   "Percentuale contributo integrativo"
         Object.Tag             =   "Percentuale contributo integrativo"
         MaxWidth        =   4
         MaxChar         =   5
         FormatMask      =   "###.00;-###.00;;"
         CanRequired     =   0   'False
         GetTextMode     =   1
      End
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_PERCCI 
         Height          =   300
         Left            =   120
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   3675
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Caption         =   "% Contr. integr."
      End
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREG 
         Height          =   300
         Left            =   10320
         TabIndex        =   28
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_PROTOCOLLO 
         Height          =   300
         Left            =   10320
         TabIndex        =   29
         Top             =   930
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         MaxChar         =   7
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   7
         CanRequired     =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "Imp. non sogg. RA"
         Height          =   225
         Left            =   6435
         TabIndex        =   66
         Top             =   2730
         Width           =   1425
      End
      Begin VB.Label Label4 
         Caption         =   "Anag. gen."
         Height          =   225
         Left            =   120
         TabIndex        =   65
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Caus. prestaz."
         Height          =   225
         Left            =   120
         TabIndex        =   64
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Data reg."
         Height          =   225
         Left            =   9540
         TabIndex        =   63
         Top             =   630
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Ditta"
         Height          =   225
         Left            =   120
         TabIndex        =   62
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Protocollo"
         Height          =   225
         Left            =   9540
         TabIndex        =   61
         Top             =   960
         Width           =   765
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   990
         TabIndex        =   1
         Top             =   300
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         MaxChar         =   6
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   6
         CanRequired     =   0   'False
         Allineamento    =   1
      End
      Begin PRJFW_EDIT.TxtEdit TXT_ANAGGEN 
         Height          =   300
         Left            =   990
         TabIndex        =   2
         Top             =   630
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         MaxChar         =   6
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   6
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPNONSOGGRA 
         Height          =   300
         Left            =   7890
         TabIndex        =   25
         Top             =   2670
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   9
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo anagrafica"
         Height          =   225
         Left            =   120
         TabIndex        =   60
         Top             =   1590
         Width           =   1155
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CBO_TIPOANAG 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         MaxChar         =   17
         IsDbField       =   0   'False
         DbCol           =   0
         CanRequired     =   0   'False
      End
   End
End
Attribute VB_Name = "FRM_RITACCINS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StrConnect       As Variant
Public CallingForm      As FRM_MAIN

Private Connessione     As ADODB.Connection
Private ClsRitenute     As CGBO_RITENUTE.CLSCG_GESTRITENUTE

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

Private Sub BUT_CALCOLAIMPORTI_Click()
    On Error GoTo Err_BUT_CALCOLAIMPORTI_Click
    
    ClsRitenute.CPInput.Sconnect = StrConnect
    ClsRitenute.CPInput.ImportoRegistrazione = TXT_IMPORTOREG.Text
    ClsRitenute.CPInput.ImportoSoggetto = TXT_IMPORTOSOGG.Text
    ClsRitenute.CPInput.ImportoIVA = TXT_IMPORTOIVA.Text
    ClsRitenute.CPInput.CausalePrestazione = TXT_CAUSPREST.Text
    ClsRitenute.CPInput.TipoAnagrafica = CBO_TIPOANAG.Text
    
    ClsRitenute.CalcolaImporti
    
    If ClsRitenute.Stato <> tsGestRitOk Then
        TXT_PROVVNONSOGG.Text = ""
        TXT_IMPONIBILERA.Text = ""
        TXT_CONTRIBINTEGR.Text = ""
        TXT_IMPONIBILERA.Text = ""
        TXT_FRANCHIGIARP.Text = ""
        TXT_IMPONIBILERP.Text = ""
        TXT_IMPORTOCPERC.Text = ""
        TXT_IMPORTOCDITTA.Text = ""
        TXT_IMPNONSOGGRA.Text = ""
        MsgBox ClsRitenute.Errore
        Exit Sub
    Else
        TXT_PROVVNONSOGG.Text = ClsRitenute.CPInput.ProvvigioniNonSoggette
        TXT_IMPONIBILERA.Text = ClsRitenute.CPInput.ImponibileSoggettoRA
        TXT_CONTRIBINTEGR.Text = ClsRitenute.CPInput.ContributoIntegrativo
        TXT_IMPORTORA.Text = ClsRitenute.CPInput.ImportoRA
        TXT_FRANCHIGIARP.Text = ClsRitenute.CPInput.FranchigiaRP
        TXT_IMPONIBILERP.Text = ClsRitenute.CPInput.ImponibileSoggettoRP
        TXT_IMPORTOCPERC.Text = ClsRitenute.CPInput.ImportoCaricoPercipiente
        TXT_IMPORTOCDITTA.Text = ClsRitenute.CPInput.ImportoCaricoDitta
        TXT_IMPNONSOGGRA.Text = ClsRitenute.CPInput.ImportoNonSoggettoRA
    End If
    
Exit Sub
Err_BUT_CALCOLAIMPORTI_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_CALCOLAIMPORTI_Click"
    Exit Sub
End Sub

Private Sub BUT_INFOANAGGEN_Click()
    On Error GoTo Err_BUT_INFOANAGGEN_Click
    
    ClsRitenute.CPInput.Ditta = TXT_DITTA.Text
    ClsRitenute.CPInput.CodiceAnagGen = TXT_ANAGGEN.Text
    
    ClsRitenute.DeterminaInfoAnagrafica
    If ClsRitenute.Stato <> tsGestRitOk Then
        CBO_TIPOANAG.Text = 0
        TXT_CAUSPREST.Text = ""
        TXT_PERCBASEIMP.Text = ""
        TXT_PERCRA.Text = ""
        TXT_CODTRIBUTO.Text = ""
        TXT_PERCINPS.Text = ""
        TXT_CODTRIBRP.Text = ""
        TXT_PERCCI.Text = ""
        TXT_PERCRIPAZ.Text = ""
        MsgBox ClsRitenute.Errore
        Exit Sub
    Else
        CBO_TIPOANAG.Text = ClsRitenute.CPInput.TipoAnagrafica
        TXT_CAUSPREST.Text = ClsRitenute.CPInput.CausalePrestazione
        TXT_PERCBASEIMP.Text = ClsRitenute.CPInput.PercBaseImponibile
        TXT_PERCRA.Text = ClsRitenute.CPInput.PercRA
        TXT_CODTRIBUTO.Text = ClsRitenute.CPInput.CodiceTributoRA
        TXT_PERCINPS.Text = ClsRitenute.CPInput.PercRP
        TXT_CODTRIBRP.Text = ClsRitenute.CPInput.CodiceTributoRP
        TXT_PERCCI.Text = ClsRitenute.CPInput.PercCI
        TXT_PERCRIPAZ.Text = ClsRitenute.CPInput.PercRipartizioneAzienda
    End If
Exit Sub
Err_BUT_INFOANAGGEN_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_INFOANAGGEN_Click"
    Exit Sub
End Sub

Private Sub BUT_INFOCAUPREST_Click()
    On Error GoTo Err_BUT_INFOCAUPREST_Click
    
    ClsRitenute.CPInput.CausalePrestazione = TXT_CAUSPREST.Text
    ClsRitenute.CPInput.Sconnect = StrConnect
    
    ClsRitenute.DeterminaInfoCausalePrestazione
    If ClsRitenute.Stato <> tsGestRitOk Then
        TXT_PERCBASEIMP.Text = ""
        TXT_PERCRA.Text = ""
        TXT_CODTRIBUTO.Text = ""
        TXT_PERCINPS.Text = ""
        TXT_CODTRIBRP.Text = ""
        TXT_PERCCI.Text = ""
        TXT_PERCRIPAZ.Text = ""
        MsgBox ClsRitenute.Errore
        Exit Sub
    Else
        TXT_PERCBASEIMP.Text = ClsRitenute.CPInput.PercBaseImponibile
        TXT_PERCRA.Text = ClsRitenute.CPInput.PercRA
        TXT_CODTRIBUTO.Text = ClsRitenute.CPInput.CodiceTributoRA
        TXT_PERCINPS.Text = ClsRitenute.CPInput.PercRP
        TXT_CODTRIBRP.Text = ClsRitenute.CPInput.CodiceTributoRP
        TXT_PERCCI.Text = ClsRitenute.CPInput.PercCI
        TXT_PERCRIPAZ.Text = ClsRitenute.CPInput.PercRipartizioneAzienda
    End If
    
Exit Sub
Err_BUT_INFOCAUPREST_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_INFOCAUPREST_Click"
    Exit Sub
End Sub

Private Sub BUT_INSERT_Click()
    On Error GoTo Err_BUT_INSERT_Click
    
    If CHK_STANDALONE.Text = 1 Then
        ClsRitenute.CPInput.MovimentoStandAlone = SiNo.Si
    Else
        ClsRitenute.CPInput.MovimentoStandAlone = SiNo.No
        ClsRitenute.CPInput.NumeroRegistrazione = TXT_NUMREG.Text
    End If
    
    ClsRitenute.CPInput.DataRegistrazione = TXT_DATAREG.Text
    ClsRitenute.CPInput.Protocollo = TXT_PROTOCOLLO.Text
    
    If CHK_BIS.Text = 1 Then
        ClsRitenute.CPInput.ProtocolloBis = SiNo.Si
    Else
        ClsRitenute.CPInput.ProtocolloBis = SiNo.No
    End If
    
    ClsRitenute.CPInput.ImportoAltreRitenute = NVL(TXT_ALTRERITENUTE.Text, 0)
    ClsRitenute.CPInput.ImportoNonSoggettoRA = NVL(TXT_IMPNONSOGGRA.Text, 0)
    
    ClsRitenute.InserisciDocumentoRitenuta
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in InserisciDocumentoRitenuta"
        Exit Sub
    End If
    
Exit Sub
Err_BUT_INSERT_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_INSERT_Click"
    Exit Sub
End Sub

Private Sub BUT_INSERTPAG_Click()
    On Error GoTo Err_BUT_INSERTPAG_Click
    
    ClsRitenute.CPInput.Ditta = TXT_DITTAPAG.Text
    ClsRitenute.CPInput.NumeroRegistrazione = TXT_NREGPERPAG.Text
    
    If CHK_STANDALONE.Text = 1 Then
        ClsRitenute.CPInput.MovimentoStandAlone = SiNo.Si
    Else
        ClsRitenute.CPInput.MovimentoStandAlone = SiNo.No
    End If
    
    ClsRitenute.CPInput.DataPagamento = TXT_DATAPAG.Text
    ClsRitenute.CPInput.MeseCompetenza = TXT_MESECOMP.Text
    ClsRitenute.CPInput.AnnoCompetenza = TXT_ANNOCOMP.Text
    ClsRitenute.CPInput.ImportoPagato = TXT_PAGATO.Text
    ClsRitenute.CPInput.ImportoAbbuono = TXT_ABBUONO.Text
    ClsRitenute.CPInput.ProvvigioniNonSoggette = TXT_PROVVNONSOG.Text
    ClsRitenute.CPInput.ContributoIntegrativo = TXT_CONTRIBINT.Text
    ClsRitenute.CPInput.ImponibileSoggettoRA = TXT_IMPON_RA.Text
    ClsRitenute.CPInput.ImportoRA = TXT_IMPORTO_RA.Text
    ClsRitenute.CPInput.ImportoAltreRitenute = TXT_IMPORTOALRIT.Text
    ClsRitenute.CPInput.ImportoNonSoggettoRA = TXT_RIMBORSI.Text
    ClsRitenute.CPInput.FranchigiaRP = TXT_FRANCHIGIA.Text
    ClsRitenute.CPInput.ImponibileSoggettoRP = TXT_IMPON_RP.Text
    ClsRitenute.CPInput.ImportoCaricoPercipiente = TXT_CONTCPERC.Text
    ClsRitenute.CPInput.ImportoCaricoDitta = TXT_CONTCDITTA.Text
    
    ClsRitenute.InserisciPagamentoRitenuta
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in InserisciPagamentoRitenuta"
        Exit Sub
    End If
    
Exit Sub
Err_BUT_INSERTPAG_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_INSERTPAG_Click"
    Exit Sub
End Sub

Private Sub BUT_PROPONIIMPORTI_Click()
    On Error GoTo Err_BUT_PROPONIIMPORTI_Click
    
    If RDB_PROP.Value = True Then
        ClsRitenute.CPInput.TipoCalcoloRitenute = Proporzionate
    Else
        ClsRitenute.CPInput.TipoCalcoloRitenute = Totali
    End If
    
    ClsRitenute.CPInput.ImportoPagato = TXT_PAGATO.Text
    ClsRitenute.CPInput.ImportoAbbuono = TXT_ABBUONO.Text
    ClsRitenute.CalcolaRitenuteDaPagare
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in CalcolaRitenuteDaPagare"
        Exit Sub
    End If
    
    TXT_PROVVNONSOG.Text = ClsRitenute.CPInput.ProvvigioniNonSoggette
    TXT_CONTRIBINT.Text = ClsRitenute.CPInput.ContributoIntegrativo
    TXT_IMPON_RA.Text = ClsRitenute.CPInput.ImponibileSoggettoRA
    TXT_IMPORTO_RA.Text = ClsRitenute.CPInput.ImportoRA
    TXT_IMPORTOALRIT.Text = ClsRitenute.CPInput.ImportoAltreRitenute
    TXT_RIMBORSI.Text = ClsRitenute.CPInput.ImportoNonSoggettoRA
    TXT_FRANCHIGIA.Text = ClsRitenute.CPInput.FranchigiaRP
    TXT_IMPON_RP.Text = ClsRitenute.CPInput.ImponibileSoggettoRP
    TXT_CONTCPERC.Text = ClsRitenute.CPInput.ImportoCaricoPercipiente
    TXT_CONTCDITTA.Text = ClsRitenute.CPInput.ImportoCaricoDitta
    
Exit Sub
Err_BUT_PROPONIIMPORTI_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_PROPONIIMPORTI_Click"
    Exit Sub
End Sub

Private Sub BUT_PROPONIPAGATO_Click()
    On Error GoTo Err_BUT_PROPONIPAGATO_Click
    
    ClsRitenute.CPInput.Ditta = TXT_DITTAPAG.Text
    ClsRitenute.CPInput.NumeroRegistrazione = TXT_NREGPERPAG.Text
    ClsRitenute.CalcolaImportoDaPagare
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in CalcolaImportoDaPagare"
        Exit Sub
    End If
    
    TXT_PAGATO.Text = ClsRitenute.CPInput.ImportoPagato
    
Exit Sub
Err_BUT_PROPONIPAGATO_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_PROPONIPAGATO_Click"
    Exit Sub
End Sub

Private Sub BUT_REGISTRAMODIFICHE_Click()
    On Error GoTo Err_BUT_REGISTRAMODIFICHE_Click
    
    ClsRitenute.RegistraModifiche
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in RegistraModifiche"
        Exit Sub
    End If
    
    If CHK_STANDALONE.Text = 1 Then
        TXT_NUMREG.Text = ClsRitenute.CPInput.NumeroRegistrazione
    End If
Exit Sub
Err_BUT_REGISTRAMODIFICHE_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_REGISTRAMODIFICHE_Click"
    Exit Sub
End Sub

Private Sub BUT_REGISTRAMODIFICHEPAG_Click()
    On Error GoTo Err_BUT_REGISTRAMODIFICHEPAG_Click
    
    ClsRitenute.RegistraModifiche
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in RegistraModifiche"
        Exit Sub
    End If
    
    If CHK_STANDALONEPAG.Text = 1 Then
        TXT_NUMREGPAG.Text = ClsRitenute.CPInput.NumRegPagamento
    End If
Exit Sub
Err_BUT_REGISTRAMODIFICHEPAG_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_REGISTRAMODIFICHEPAG_Click"
    Exit Sub
End Sub

Private Sub CHK_STANDALONE_Click()
    On Error GoTo Err_CHK_STANDALONE_Click
    
    If CHK_STANDALONE.Text = 1 Then
        TXT_DATAREG.Enabled = True
        TXT_PROTOCOLLO.Enabled = True
        CHK_BIS.Enabled = True
        TXT_NUMREG.Text = ""
        TXT_NUMREG.Enabled = False
    Else
        TXT_DATAREG.Text = ""
        TXT_DATAREG.Enabled = False
        TXT_PROTOCOLLO.Text = ""
        TXT_PROTOCOLLO.Enabled = False
        CHK_BIS.Text = 0
        CHK_BIS.Enabled = False
        TXT_NUMREG.Enabled = True
    End If
    
Exit Sub
Err_CHK_STANDALONE_Click:
    MsgBox Err.Number & " - " & Err.Description, , "CHK_STANDALONE_Click"
    Exit Sub
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Form_Load
    
    Set ClsRitenute = New CGBO_RITENUTE.CLSCG_GESTRITENUTE
    ClsRitenute.CPInput.Sconnect = StrConnect
    
    With CBO_TIPOANAG
        .AddItemData "No Ritenuta Acconto", 0
        .AddItemData "Professionista", 1
        .AddItemData "Rappresentante", 2
    End With
    
    CHK_STANDALONE.Text = 1
Exit Sub
Err_Form_Load:
    MsgBox Err.Number & " - " & Err.Description, , "Form_Load"
    Exit Sub
End Sub

Private Sub TXT_CAUSPREST_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    Dim Pcls_Lookup As COBO_LOOKUPDECODE.CLSCO_LOOKUP
    
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Lookup = New COBO_LOOKUPDECODE.CLSCO_LOOKUP
    
    Pcls_Lookup.CausaliPrestazione
    
    str_SQL = Pcls_Lookup.StringaSQL
    Arr_Fields = Pcls_Lookup.ArrayFields
    Str_Caption = Pcls_Lookup.Titolo
    Str_Connect = StrConnect
    
    Set Pcls_Lookup = Nothing
    Err.Clear
End Sub

Private Sub TXT_CODTRIBRP_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    Dim Pcls_Lookup As CGBO_LOOKUPDECODE.CLSCG_LOOKUP
    
    On Error GoTo ErrTrap
    
    Cancel = False
    
    Set Pcls_Lookup = New CGBO_LOOKUPDECODE.CLSCG_LOOKUP
    
    Call Pcls_Lookup.Tributo(2)
    
    str_SQL = Pcls_Lookup.StringaSQL
    Arr_Fields = Pcls_Lookup.ColonneLookup
    Str_Caption = Pcls_Lookup.Caption
    Str_Connect = StrConnect
    
    TXT_CODTRIBUTO.IDLookup = Pcls_Lookup.IDLookup
    
    Set Pcls_Lookup = Nothing

Exit Sub
ErrTrap:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CODTRIBRP_StartLookup"
    Exit Sub
End Sub

Private Sub TXT_CODTRIBUTO_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    Dim Pcls_Lookup As CGBO_LOOKUPDECODE.CLSCG_LOOKUP
    
    On Error GoTo ErrTrap
    
    Cancel = False
    
    Set Pcls_Lookup = New CGBO_LOOKUPDECODE.CLSCG_LOOKUP
    
    Call Pcls_Lookup.Tributo(1)
    
    str_SQL = Pcls_Lookup.StringaSQL
    Arr_Fields = Pcls_Lookup.ColonneLookup
    Str_Caption = Pcls_Lookup.Caption
    Str_Connect = StrConnect
    
    TXT_CODTRIBUTO.IDLookup = Pcls_Lookup.IDLookup
    
    Set Pcls_Lookup = Nothing

Exit Sub
ErrTrap:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CODTRIBUTO_StartLookup"
    Exit Sub
End Sub

