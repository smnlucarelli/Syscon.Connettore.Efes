VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.7#0"; "TMS_COMBOBOX.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.6#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.5#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.4#0"; "TMS_EDITDATE.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F53BE214-7AC6-11D0-9B0E-006097A80EFD}#6.6#0"; "TMS_LABEL.ocx"
Object = "{0EF4EAD5-2617-11D2-A1C0-0060082875F9}#5.7#0"; "TMS_CHECKBOX.ocx"
Object = "{C99C525C-61F9-11D2-AE21-00A0244C5B50}#3.7#0"; "TMS_EDITNUMM.ocx"
Begin VB.Form FRM_RITACCUPD 
   Caption         =   "Gestione ritenute acconto - modifica"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   Icon            =   "FRM_RITACCUPD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab TAB_RIT 
      Height          =   4935
      Left            =   30
      TabIndex        =   31
      Top             =   750
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tabella CG54_DOCRITAC"
      TabPicture(0)   =   "FRM_RITACCUPD.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GRID_CG54"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tabella CG55_PAGRITAC"
      TabPicture(1)   =   "FRM_RITACCUPD.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "GRID_CG55"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tabella CG41_PRIMANOTA"
      TabPicture(2)   =   "FRM_RITACCUPD.frx":27DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GRID_CG41PAG"
      Tab(2).Control(1)=   "GRID_CG41RIT"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Modifica pagamento ritenuta:"
         Height          =   2985
         Left            =   -74940
         TabIndex        =   43
         Top             =   1890
         Width           =   9285
         Begin VB.Frame Frame4 
            Caption         =   "Dati versamento ritenute d'acconto"
            Height          =   2325
            Left            =   4050
            TabIndex        =   49
            Top             =   180
            Width           =   4725
            Begin PRJFW_EDITNUMM.TxtEditNumM TXT_ABIRA 
               Height          =   300
               Left            =   930
               TabIndex        =   20
               Top             =   930
               Width           =   1290
               _ExtentX        =   2302
               _ExtentY        =   529
               IsLookup        =   -1  'True
               TxtGestione     =   "Banche e agenzie"
               IsGestione      =   -1  'True
               IsCalculator    =   0   'False
               MaxChar         =   5
               DBField         =   "CG55_CODABIRA_CG12"
               Caption         =   "Codice ABI rit. acc."
               Object.Tag             =   "Codice ABI rit. acc."
               MaxWidth        =   4
               IsDecode        =   -1  'True
            End
            Begin PRJFW_EDITNUMM.TxtEditNumM TXT_CABRA 
               Height          =   300
               Left            =   930
               TabIndex        =   21
               Top             =   1260
               Width           =   1290
               _ExtentX        =   2302
               _ExtentY        =   529
               IsLookup        =   -1  'True
               TxtGestione     =   "Banche e Agenzie"
               IsGestione      =   -1  'True
               IsCalculator    =   0   'False
               MaxChar         =   5
               DBField         =   "CG55_CODCABRA_CG13"
               Caption         =   "Codice  CAB rit. acc."
               Object.Tag             =   "Codice  CAB rit. acc."
               MaxWidth        =   4
               IsDecode        =   -1  'True
            End
            Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL23 
               Height          =   300
               Left            =   2280
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   1950
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   529
               Caption         =   "C/C postale"
            End
            Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL5 
               Height          =   300
               Left            =   120
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   600
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   529
               Caption         =   "A mezzo"
            End
            Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL12 
               Height          =   300
               Left            =   120
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   960
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   529
               Caption         =   "Codice ABI"
            End
            Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL13 
               Height          =   300
               Left            =   120
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   1290
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   529
               Caption         =   "Codice CAB"
            End
            Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL4 
               Height          =   300
               Left            =   120
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1620
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   529
               Caption         =   "Serie"
            End
            Begin PRJFW_EDIT.TxtEdit TXT_SERIERA 
               Height          =   300
               Left            =   930
               TabIndex        =   22
               ToolTipText     =   "Serie di versamento tramite esattoria"
               Top             =   1590
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   529
               MaxChar         =   4
               Numerico        =   0   'False
               Carattere       =   0   'False
               DBField         =   "CG55_SERIERA"
               Caption         =   "Serie di versamento tramite esattoria"
               Object.Tag             =   "Serie di versamento tramite esattoria"
               MaxWidth        =   5
            End
            Begin PRJFW_EDIT.TxtEdit TXT_CCPOSTALERA 
               Height          =   300
               Left            =   3030
               TabIndex        =   24
               ToolTipText     =   "Numero C/C postale"
               Top             =   1920
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   529
               MaxChar         =   12
               Numerico        =   0   'False
               Carattere       =   0   'False
               DBField         =   "CG55_CCPOSTALERA"
               Caption         =   "Numero C/C postale"
               Object.Tag             =   "Numero C/C postale"
               MaxWidth        =   11
            End
            Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL22 
               Height          =   300
               Left            =   2280
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   1620
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   529
               Caption         =   "Quietanza"
            End
            Begin PRJFW_EDIT.TxtEdit TXT_QUIETANZARA 
               Height          =   300
               Left            =   3030
               TabIndex        =   23
               ToolTipText     =   "Quietanza di versamento tramite esattoria"
               Top             =   1590
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   529
               MaxChar         =   12
               Numerico        =   0   'False
               Carattere       =   0   'False
               DBField         =   "CG55_QUIETANZARA"
               Caption         =   "Quietanza di versamento tramite esattoria"
               Object.Tag             =   "Quietanza di versamento tramite esattoria"
               MaxWidth        =   11
            End
            Begin PRJFW_EDIT.TxtEdit TXT_DESCRABIRA 
               Height          =   300
               Left            =   2250
               TabIndex        =   52
               Top             =   930
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   529
               Enabled         =   0   'False
               MaxChar         =   40
               Numerico        =   0   'False
               Carattere       =   0   'False
               IsDbField       =   0   'False
               MaxWidth        =   19
               CanRequired     =   0   'False
            End
            Begin PRJFW_EDIT.TxtEdit TXT_DESCRCABRA 
               Height          =   300
               Left            =   2250
               TabIndex        =   51
               Top             =   1260
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   529
               Enabled         =   0   'False
               MaxChar         =   40
               Numerico        =   0   'False
               Carattere       =   0   'False
               IsDbField       =   0   'False
               MaxWidth        =   19
               CanRequired     =   0   'False
            End
            Begin PRJFW_EDITNUM.TxtEditNum TXT_ANNOVERSRA 
               Height          =   300
               Left            =   1410
               TabIndex        =   18
               Top             =   240
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   529
               DBField         =   "CG55_ANNOVERS"
               Caption         =   "Anno versamento ritenuta acconto"
               Object.Tag             =   "Anno versamento ritenuta acconto"
               MaxWidth        =   4
               MaxChar         =   4
               CanRequired     =   0   'False
            End
            Begin PRJFW_EDITNUM.TxtEditNum TXT_MESEVERSRA 
               Height          =   300
               Left            =   930
               TabIndex        =   17
               Top             =   240
               Width           =   510
               _ExtentX        =   900
               _ExtentY        =   529
               DBField         =   "CG55_MESEVERS"
               Caption         =   "Mese versamento ritenuta acconto"
               Object.Tag             =   "Mese versamento ritenuta acconto"
               MaxWidth        =   2
               MaxChar         =   2
            End
            Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL9 
               Height          =   300
               Left            =   120
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   270
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   529
               Caption         =   "Mese/Anno"
            End
            Begin PRJFW_COMBOBOX.TMS_COMBO CMB_TIPOVERSRA 
               Height          =   315
               Left            =   1110
               TabIndex        =   19
               Top             =   570
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   556
               MaxChar         =   18
               Default         =   "0"
               DBField         =   "CG55_INDTIPOVERSRA"
               DbCol           =   0
               Caption         =   "Modalità di versamento ritenute d'acconto"
               Object.Tag             =   "Modalità di versamento ritenute d'acconto"
               CanRequired     =   0   'False
            End
         End
         Begin VB.CommandButton BUT_MODIFICAPAG 
            Caption         =   "Modifica pagamento"
            Height          =   345
            Left            =   7530
            Picture         =   "FRM_RITACCUPD.frx":27F6
            TabIndex        =   26
            Top             =   2550
            Width           =   1665
         End
         Begin VB.Label Label3 
            Caption         =   "Data pagamento"
            Height          =   225
            Left            =   120
            TabIndex        =   59
            Top             =   870
            Width           =   1335
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAPAG 
            Height          =   300
            Left            =   1695
            TabIndex        =   12
            Top             =   810
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            IsCalendario    =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin PRJFW_CHECKBOX.TMS_CHECKBOX CHK_FLGSTAMPRA 
            Height          =   300
            Left            =   4020
            TabIndex        =   25
            Top             =   2520
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            Default         =   "0"
            DBField         =   "CG55_FLGSTAMPRA"
            Caption         =   "Certificazione stampata"
            Object.Tag             =   "Ritenuta d'acconto stampata"
         End
         Begin VB.Label Label2 
            Caption         =   "Num. reg. pagamento"
            Height          =   225
            Left            =   120
            TabIndex        =   48
            Top             =   300
            Width           =   1545
         End
         Begin PRJFW_EDIT.TxtEdit TXT_NUMREGPAG 
            Height          =   300
            Left            =   1695
            TabIndex        =   11
            Top             =   270
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
         Begin VB.Label Label23 
            Caption         =   "Abbuono"
            Height          =   225
            Left            =   120
            TabIndex        =   47
            Top             =   1530
            Width           =   735
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_ABBUONO 
            Height          =   300
            Left            =   1695
            TabIndex        =   14
            Top             =   1470
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            IsDbField       =   0   'False
            Caption         =   "Importo abbuono"
            Object.Tag             =   "Importo abbuono"
            MaxWidth        =   11
            MaxChar         =   13
            FormatMask      =   "##,###,###,##0.00"
            CanRequired     =   0   'False
         End
         Begin VB.Label Label9 
            Caption         =   "Pagato"
            Height          =   225
            Left            =   120
            TabIndex        =   46
            Top             =   1200
            Width           =   645
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_PAGATO 
            Height          =   300
            Left            =   1695
            TabIndex        =   13
            Top             =   1140
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            IsDbField       =   0   'False
            Caption         =   "Importo pagato"
            Object.Tag             =   "Importo pagato"
            MaxWidth        =   11
            MaxChar         =   13
            FormatMask      =   "##,###,###,##0.00"
            CanRequired     =   0   'False
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL3 
            Height          =   300
            Left            =   120
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1830
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   529
            Caption         =   "Imponibile sogg. R.A."
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL1 
            Height          =   195
            Left            =   120
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            Caption         =   "Importo R.A."
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPSOGGRA 
            Height          =   300
            Left            =   1695
            TabIndex        =   15
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            IsDbField       =   0   'False
            Caption         =   "Imponibile soggetto RA"
            Object.Tag             =   "Imponibile soggetto R.A."
            MaxWidth        =   11
            MaxChar         =   13
            FormatMask      =   "##,###,###,##0.00"
            CanRequired     =   0   'False
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTORA 
            Height          =   300
            Left            =   1695
            TabIndex        =   16
            Top             =   2130
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            IsDbField       =   0   'False
            Caption         =   "Importo RA"
            Object.Tag             =   "Importo ritenuta d'acconto"
            MaxWidth        =   11
            MaxChar         =   13
            FormatMask      =   "##,###,###,##0.00"
            CanRequired     =   0   'False
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Modifica testata ritenuta:"
         Height          =   2055
         Left            =   60
         TabIndex        =   36
         Top             =   1560
         Width           =   6885
         Begin VB.CommandButton BUT_MODIFICARIT 
            Caption         =   "Modifica ritenuta"
            Height          =   345
            Left            =   5280
            Picture         =   "FRM_RITACCUPD.frx":2940
            TabIndex        =   10
            Top             =   1620
            Width           =   1485
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL32 
            Height          =   300
            Left            =   3090
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   270
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            Caption         =   "Imponibile soggetto R.A."
         End
         Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREG 
            Height          =   300
            Left            =   990
            TabIndex        =   3
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            IsCalendario    =   0   'False
            IsDbField       =   0   'False
            CanRequired     =   0   'False
         End
         Begin VB.Label Label17 
            Caption         =   "Data reg."
            Height          =   225
            Left            =   120
            TabIndex        =   42
            Top             =   270
            Width           =   1095
         End
         Begin PRJFW_CHECKBOX.TMS_CHECKBOX CHK_BIS 
            Height          =   300
            Left            =   1920
            TabIndex        =   5
            Top             =   600
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   529
            IsDbField       =   0   'False
            Caption         =   "Bis"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_PROTOCOLLO 
            Height          =   300
            Left            =   990
            TabIndex        =   4
            Top             =   570
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
         Begin VB.Label Label12 
            Caption         =   "Protocollo"
            Height          =   225
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   765
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_IMPORTO_RA 
            Height          =   300
            Left            =   4770
            TabIndex        =   7
            Top             =   570
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
            Left            =   4770
            TabIndex        =   6
            Top             =   240
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
            Left            =   4770
            TabIndex        =   9
            Top             =   1230
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
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL34 
            Height          =   195
            Left            =   3090
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            Caption         =   "Importo R.A."
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL42 
            Height          =   300
            Left            =   3090
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1260
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            Caption         =   "Importo non soggetto"
         End
         Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL2 
            Height          =   300
            Left            =   3090
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   930
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            Caption         =   "Contr. integr."
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_CONTRIBINT 
            Height          =   300
            Left            =   4770
            TabIndex        =   8
            Top             =   900
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   529
            IsDbField       =   0   'False
            Caption         =   "Contributo integrativo"
            Object.Tag             =   "Importo contributo integrativo"
            MaxWidth        =   11
            MaxChar         =   13
            FormatMask      =   "##,###,###,##0.00"
         End
      End
      Begin MSDataGridLib.DataGrid GRID_CG54 
         Height          =   1125
         Left            =   30
         TabIndex        =   32
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   1984
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
      Begin MSDataGridLib.DataGrid GRID_CG41RIT 
         Height          =   1545
         Left            =   -74970
         TabIndex        =   33
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2725
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
      Begin MSDataGridLib.DataGrid GRID_CG41PAG 
         Height          =   2895
         Left            =   -74970
         TabIndex        =   34
         Top             =   1950
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5106
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
      Begin MSDataGridLib.DataGrid GRID_CG55 
         Height          =   1515
         Left            =   -74970
         TabIndex        =   35
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
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
   Begin VB.Frame Frame1 
      Caption         =   "Dati relativi alla ritenuta:"
      Height          =   645
      Left            =   30
      TabIndex        =   27
      Top             =   30
      Width           =   7005
      Begin VB.CommandButton BUT_QUERY 
         Caption         =   "Ricerca mov. ritenuta"
         Height          =   345
         Left            =   4980
         Picture         =   "FRM_RITACCUPD.frx":2A8A
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Num. reg. ritenuta"
         Height          =   225
         Left            =   1845
         TabIndex        =   30
         Top             =   270
         Width           =   1335
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
         Height          =   300
         Left            =   3240
         TabIndex        =   1
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
      Begin VB.Label Label1 
         Caption         =   "Ditta"
         Height          =   225
         Left            =   120
         TabIndex        =   29
         Top             =   270
         Width           =   495
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   660
         TabIndex        =   0
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
   End
   Begin VB.CommandButton BUT_REGISTRAMODIFICHE 
      Caption         =   "Registra Modifiche"
      Height          =   345
      Left            =   9000
      Picture         =   "FRM_RITACCUPD.frx":2BD4
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5730
      Width           =   1785
   End
End
Attribute VB_Name = "FRM_RITACCUPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StrConnect       As Variant
Public CallingForm      As FRM_MAIN

Private Connessione     As ADODB.Connection
Private ClsRitenute     As CGBO_RITENUTE.CLSCG_GESTRITENUTE

Private Pcls_Decode     As COBO_LOOKUPDECODE.CLSCO_DECODE
Private Pcls_Lookup     As COBO_LOOKUPDECODE.CLSCO_LOOKUP

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

Private Sub BUT_MODIFICAPAG_Click()
    Dim RS      As ADODB.Recordset
    
    On Error GoTo Err_BUT_MODIFICAPAG_Click
    
    ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitNonDefinito
    
    Set RS = GRID_CG55.DataSource
    
    If TXT_DATAPAG.Text <> RS.Fields("CG55_DATAPAG").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitDataPagamento
        ClsRitenute.CPInput.DataPagamento = TXT_DATAPAG.Text
    End If
    If TXT_PAGATO.Text <> RS.Fields("CG55_IMPORTOPAG").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitImportoPagato
        ClsRitenute.CPInput.ImportoPagato = TXT_PAGATO.Text
    End If
    If TXT_ABBUONO.Text <> RS.Fields("CG55_IMPORTOABB").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitImportoAbbuono
        ClsRitenute.CPInput.ImportoAbbuono = TXT_ABBUONO.Text
    End If
    If TXT_IMPSOGGRA.Text <> RS.Fields("CG55_IMPONIBILERA").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitImponibileSoggettoRA
        ClsRitenute.CPInput.ImponibileSoggettoRA = TXT_IMPSOGGRA.Text
    End If
    If TXT_IMPORTORA.Text <> RS.Fields("CG55_IMPORTORA").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitImportoRA
        ClsRitenute.CPInput.ImportoRA = TXT_IMPORTORA.Text
    End If
    If TXT_MESEVERSRA.Text <> RS.Fields("CG55_MESEVERS").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitMeseVersRA
        ClsRitenute.CPInput.MeseVersRA = TXT_MESEVERSRA.Text
    End If
    If TXT_ANNOVERSRA.Text <> RS.Fields("CG55_ANNOVERS").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitAnnoVersRA
        ClsRitenute.CPInput.AnnoVersRA = TXT_ANNOVERSRA.Text
    End If
    If CMB_TIPOVERSRA.Text <> RS.Fields("CG55_INDTIPOVERSRA").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitTipoVersamentoRA
        ClsRitenute.CPInput.TipoVersamentoRA = CMB_TIPOVERSRA.Text
    End If
    If TXT_ABIRA.Text <> RS.Fields("CG55_CODABIRA_CG12").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitCodiceABIRA
        ClsRitenute.CPInput.CodiceABIRA = TXT_ABIRA.Text
    End If
    If TXT_CABRA.Text <> RS.Fields("CG55_CODCABRA_CG13").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitCodiceCABRA
        ClsRitenute.CPInput.CodiceCABRA = TXT_CABRA.Text
    End If
    If TXT_SERIERA.Text <> RS.Fields("CG55_SERIERA").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitNumeroSerieRA
        ClsRitenute.CPInput.NumeroSerieRA = TXT_SERIERA.Text
    End If
    If TXT_QUIETANZARA.Text <> RS.Fields("CG55_QUIETANZARA").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitNumeroQuietanzaRA
        ClsRitenute.CPInput.NumeroQuietanzaRA = TXT_QUIETANZARA.Text
    End If
    If TXT_CCPOSTALERA.Text <> RS.Fields("CG55_CCPOSTALERA").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitNumeroCCPostaleRA
        ClsRitenute.CPInput.NumeroCCPostaleRA = TXT_CCPOSTALERA.Text
    End If
    If CHK_FLGSTAMPRA.Text <> RS.Fields("CG55_FLGSTAMPRA").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiPagRit CampoPagRitStampaRitenutaAcconto
        If CHK_FLGSTAMPRA.Text = 1 Then
            ClsRitenute.CPInput.StampaRitenutaAcconto = SiNo.Si
        Else
            ClsRitenute.CPInput.StampaRitenutaAcconto = SiNo.No
        End If
    End If
    
    ClsRitenute.CPInput.NumRegPagamento = TXT_NUMREGPAG.Text
    ClsRitenute.ModificaPagamentoRitenuta
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in ModificaPagamentoRitenuta"
        Exit Sub
    End If
    
Exit Sub
Err_BUT_MODIFICAPAG_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_MODIFICAPAG_Click"
    Exit Sub
End Sub

Private Sub BUT_MODIFICARIT_Click()
    On Error GoTo Err_BUT_MODIFICARIT_Click
    
    ClsRitenute.CPInput.AggiungiVoceAdElencoCampiDocRit CampoDocRitNonDefinito
    
    If TXT_DATAREG.Text <> ClsRitenute.RecSetCG41Ritenuta.Fields("CG41_DATAREG").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiDocRit CampoDocRitDataRegistrazione
        ClsRitenute.CPInput.DataRegistrazione = TXT_DATAREG.Text
    End If
    If TXT_PROTOCOLLO.Text <> ClsRitenute.RecSetCG41Ritenuta.Fields("CG41_NUMDOC").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiDocRit CampoDocRitProtocollo
        ClsRitenute.CPInput.Protocollo = TXT_PROTOCOLLO.Text
    End If
    If CHK_BIS.Text <> ClsRitenute.RecSetCG41Ritenuta.Fields("CG41_FLGDOCBIS").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiDocRit CampoDocRitProtocolloBis
        If CHK_BIS.Text = 1 Then
            ClsRitenute.CPInput.ProtocolloBis = SiNo.Si
        Else
            ClsRitenute.CPInput.ProtocolloBis = SiNo.No
        End If
    End If
    If TXT_IMPON_RA.Text <> ClsRitenute.RecSetCG54.Fields("CG54_IMPONIBILERA").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiDocRit CampoDocRitImponibileSoggettoRA
        ClsRitenute.CPInput.ImponibileSoggettoRA = TXT_IMPON_RA.Text
    End If
    If TXT_IMPORTO_RA.Text <> ClsRitenute.RecSetCG54.Fields("CG54_IMPORTORA").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiDocRit CampoDocRitImportoRA
        ClsRitenute.CPInput.ImportoRA = TXT_IMPORTO_RA.Text
    End If
    If TXT_CONTRIBINT.Text <> ClsRitenute.RecSetCG54.Fields("CG54_CONTRIBINT").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiDocRit CampoDocRitContributoIntegrativo
        ClsRitenute.CPInput.ContributoIntegrativo = TXT_CONTRIBINT.Text
    End If
    If TXT_RIMBORSI.Text <> ClsRitenute.RecSetCG54.Fields("CG54_RIMBORSI").Value Then
        ClsRitenute.CPInput.AggiungiVoceAdElencoCampiDocRit CampoDocRitImportoNonSoggettoRA
        ClsRitenute.CPInput.ImportoNonSoggettoRA = TXT_RIMBORSI.Text
    End If
    
    ClsRitenute.ModificaDocumentoRitenuta
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in ModificaDocumentoRitenuta"
        Exit Sub
    End If
    
Exit Sub
Err_BUT_MODIFICARIT_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_MODIFICARIT_Click"
    Exit Sub
End Sub

Private Sub BUT_QUERY_Click()
    On Error GoTo Err_BUT_QUERY_Click
    
    ClsRitenute.CPInput.Ditta = TXT_DITTA.Text
    ClsRitenute.CPInput.NumeroRegistrazione = TXT_NUMREG.Text
    ClsRitenute.CPInput.ForzaCancellazioneRitenutaDaPrimaNota = True
    ClsRitenute.CPInput.ForzaCancellazionePagRitDaPrimaNota = True
    
    ClsRitenute.ModificaMovimentoRitenuta
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in CancellaRitenuta"
        Exit Sub
    End If
    
    Set GRID_CG54.DataSource = ClsRitenute.RecSetCG54
    Set GRID_CG55.DataSource = ClsRitenute.RecSetCG55
    
    Set GRID_CG41RIT.DataSource = ClsRitenute.RecSetCG41Ritenuta
    Set GRID_CG41PAG.DataSource = ClsRitenute.RecSetCG41Pagamenti
    
    TXT_DATAREG.Text = ClsRitenute.RecSetCG41Ritenuta.Fields("CG41_DATAREG").Value
    TXT_PROTOCOLLO.Text = ClsRitenute.RecSetCG41Ritenuta.Fields("CG41_NUMDOC").Value
    CHK_BIS.Text = ClsRitenute.RecSetCG41Ritenuta.Fields("CG41_FLGDOCBIS").Value
    
    TXT_IMPON_RA.Text = ClsRitenute.RecSetCG54.Fields("CG54_IMPONIBILERA").Value
    TXT_IMPORTO_RA.Text = ClsRitenute.RecSetCG54.Fields("CG54_IMPORTORA").Value
    TXT_CONTRIBINT.Text = ClsRitenute.RecSetCG54.Fields("CG54_CONTRIBINT").Value
    TXT_RIMBORSI.Text = ClsRitenute.RecSetCG54.Fields("CG54_RIMBORSI").Value
Exit Sub
Err_BUT_QUERY_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_QUERY_Click"
    Exit Sub
End Sub

Private Sub BUT_REGISTRAMODIFICHE_Click()
    On Error GoTo Err_BUT_REGISTRAMODIFICHE_Click
    
    ClsRitenute.RegistraModifiche
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in RegistraModifiche"
        Exit Sub
    End If
Exit Sub
Err_BUT_REGISTRAMODIFICHE_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_REGISTRAMODIFICHE_Click"
    Exit Sub
End Sub

Private Sub Form_Activate()
    On Error GoTo Err_Form_Activate
    
    CMB_TIPOVERSRA.EraseCombo
    CMB_TIPOVERSRA.AddItemData "Non effettuato", 0
    CMB_TIPOVERSRA.AddItemData "Esattoria", 1
    CMB_TIPOVERSRA.AddItemData "C/C postale", 2
    CMB_TIPOVERSRA.AddItemData "Banca", 3
    CMB_TIPOVERSRA.AutoOpen = False
    CMB_TIPOVERSRA.Text = 0
Exit Sub
Err_Form_Activate:
    MsgBox Err.Number & " - " & Err.Description, , "Form_Activate"
    Exit Sub
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
    
    Set ClsRitenute = New CGBO_RITENUTE.CLSCG_GESTRITENUTE
    ClsRitenute.CPInput.SConnect = StrConnect
    TAB_RIT.Tab = 0
    
    Set Pcls_Decode = New COBO_LOOKUPDECODE.CLSCO_DECODE
    Set Pcls_Lookup = New COBO_LOOKUPDECODE.CLSCO_LOOKUP
    
    Set TXT_ABIRA.ActiveInterface = CallingForm.ActiveInterface
    Set TXT_CABRA.ActiveInterface = CallingForm.ActiveInterface
    
    Set CallingForm.ActiveInterface.Connection = Connessione
Exit Sub
Err_Form_Load:
    MsgBox Err.Number & " - " & Err.Description, , "Form_Load"
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Set Pcls_Decode = Nothing
    Set Pcls_Lookup = Nothing
    
    Err.Clear
End Sub

Private Sub GRID_CG55_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim RS      As ADODB.Recordset
    
    On Error GoTo Err_GRID_CG55_RowColChange
    
    Set RS = GRID_CG55.DataSource
    TXT_NUMREGPAG.Text = RS.Fields("CG55_REGPN_CG41").Value
    TXT_DATAPAG.Text = RS.Fields("CG55_DATAPAG").Value
    TXT_PAGATO.Text = RS.Fields("CG55_IMPORTOPAG").Value
    TXT_ABBUONO.Text = RS.Fields("CG55_IMPORTOABB").Value
    TXT_IMPSOGGRA.Text = RS.Fields("CG55_IMPONIBILERA").Value
    TXT_IMPORTORA.Text = RS.Fields("CG55_IMPORTORA").Value
    TXT_MESEVERSRA.Text = RS.Fields("CG55_MESEVERS").Value
    TXT_ANNOVERSRA.Text = RS.Fields("CG55_ANNOVERS").Value
    CMB_TIPOVERSRA.Text = RS.Fields("CG55_INDTIPOVERSRA").Value
    TXT_ABIRA.Text = RS.Fields("CG55_CODABIRA_CG12").Value
    TXT_CABRA.Text = RS.Fields("CG55_CODCABRA_CG13").Value
    TXT_SERIERA.Text = RS.Fields("CG55_SERIERA").Value
    TXT_QUIETANZARA.Text = RS.Fields("CG55_QUIETANZARA").Value
    TXT_CCPOSTALERA.Text = RS.Fields("CG55_CCPOSTALERA").Value
    CHK_FLGSTAMPRA.Text = RS.Fields("CG55_FLGSTAMPRA").Value
    
Exit Sub
Err_GRID_CG55_RowColChange:
    MsgBox Err.Number & " - " & Err.Description, , "GRID_CG55_RowColChange"
    Exit Sub
End Sub

Private Sub TXT_ABIRA_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error GoTo Err_TXT_ABIRA_StartDecode
    
    Set Pcls_Decode.CampoDecodifica = TXT_DESCRABIRA
    Pcls_Decode.Banca TXT_ABIRA.Text
    
    str_SQL = Pcls_Decode.StringaSQL
    Arr_Fields = Pcls_Decode.ArrayFields
    Str_Connect = StrConnect
    
Exit Sub
Err_TXT_ABIRA_StartDecode:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_ABIRA_StartDecode"
    Exit Sub
End Sub

Private Sub TXT_ABIRA_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_ABIRA_StartLookup
    
    Call Pcls_Lookup.Banche
    
    str_SQL = Pcls_Lookup.StringaSQL
    Arr_Fields = Pcls_Lookup.ArrayFields
    Str_Caption = Pcls_Lookup.Titolo
    Str_Connect = StrConnect
    
    TXT_ABIRA.IDLookup = Pcls_Lookup.IDLookup
Exit Sub
Err_TXT_ABIRA_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_ABIRA_StartLookup"
    Exit Sub
End Sub

Private Sub TXT_CABRA_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error GoTo Err_TXT_CABRA_StartDecode
    
    Set Pcls_Decode.CampoDecodifica = TXT_DESCRCABRA
    Pcls_Decode.Agenzia TXT_ABIRA.Text, TXT_CABRA.Text
    
    str_SQL = Pcls_Decode.StringaSQL
    Arr_Fields = Pcls_Decode.ArrayFields
    Str_Connect = StrConnect
Exit Sub
Err_TXT_CABRA_StartDecode:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CABRA_StartDecode"
    Exit Sub
End Sub

Private Sub TXT_CABRA_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    On Error GoTo Err_TXT_CABRA_StartLookup
    
    Pcls_Lookup.Agenzia TXT_ABIRA.Text
    
    str_SQL = Pcls_Lookup.StringaSQL
    Arr_Fields = Pcls_Lookup.ArrayFields
    Str_Caption = Pcls_Lookup.Titolo
    Str_Connect = StrConnect
    
    TXT_CABRA.IDLookup = Pcls_Lookup.IDLookup
Exit Sub
Err_TXT_CABRA_StartLookup:
    MsgBox Err.Number & " - " & Err.Description, , "TXT_CABRA_StartLookup"
    Exit Sub
End Sub
