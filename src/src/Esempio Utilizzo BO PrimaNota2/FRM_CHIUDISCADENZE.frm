VERSION 5.00
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.7#0"; "TMS_EDITM.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.6#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.5#0"; "TMS_EDITNUM.ocx"
Object = "{F53BE214-7AC6-11D0-9B0E-006097A80EFD}#6.6#0"; "TMS_LABEL.ocx"
Object = "{0EF4EAD5-2617-11D2-A1C0-0060082875F9}#5.7#0"; "TMS_CHECKBOX.ocx"
Object = "{F2DC983F-61F7-11D2-AE21-00A0244C5B50}#3.6#0"; "TMS_EDITDATEM.ocx"
Begin VB.Form FRM_CHIUDISCADENZE 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BUT_ESEGUI 
      Caption         =   "Esegui"
      Height          =   405
      Left            =   7170
      Picture         =   "FRM_CHIUDISCADENZE.frx":0000
      TabIndex        =   16
      Top             =   4050
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pagamento / incasso scadenze:"
      Height          =   3915
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8145
      Begin VB.Frame Frame2 
         Caption         =   "Pagato:"
         Height          =   1275
         Left            =   2910
         TabIndex        =   25
         Top             =   1440
         Width           =   5175
         Begin PRJFW_EDITM.TXT_EDITM TXT_CAUSALEPAG 
            Height          =   300
            Left            =   1410
            TabIndex        =   11
            Top             =   540
            Width           =   1020
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
         Begin PRJFW_EDIT.TxtEdit TXT_DESCRCAUSALE 
            Height          =   300
            Left            =   1410
            TabIndex        =   12
            Top             =   870
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   529
            MaxChar         =   200
            Numerico        =   0   'False
            Carattere       =   0   'False
            IsDbField       =   0   'False
            MaxWidth        =   30
            CanRequired     =   0   'False
         End
         Begin VB.Label Label6 
            Caption         =   "Causale"
            Height          =   225
            Left            =   150
            TabIndex        =   28
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Importo pagato"
            Height          =   225
            Left            =   150
            TabIndex        =   27
            Top             =   270
            Width           =   1125
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_PAGATO 
            Height          =   300
            Left            =   1410
            TabIndex        =   10
            Top             =   210
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   529
            IsDbField       =   0   'False
            MaxChar         =   13
            FormatMask      =   "##,###,###,##0.00"
            CanRequired     =   0   'False
         End
         Begin VB.Label Label10 
            Caption         =   "Descr. causale"
            Height          =   225
            Left            =   150
            TabIndex        =   26
            Top             =   930
            Width           =   1245
         End
      End
      Begin PRJFW_EDITM.TXT_EDITM TXT_VALUTA 
         Height          =   300
         Left            =   4320
         TabIndex        =   8
         Top             =   750
         Width           =   1020
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
      Begin PRJFW_EDITDATEM.EditDateM TXT_DATAREGPAG 
         Height          =   300
         Left            =   4320
         TabIndex        =   15
         Top             =   3450
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         IsCalendario    =   0   'False
         MaxChar         =   8
         IsDbField       =   0   'False
         Caption         =   "Data pagamento"
         Object.Tag             =   "Data pagamento fattura con ritenute"
         Formato         =   "Medium Date"
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITM.TXT_EDITM TXT_SEZPART 
         Height          =   300
         Left            =   1170
         TabIndex        =   5
         Top             =   2070
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   529
         IsLookup        =   0   'False
         DisplayFormat   =   "Maiuscolo"
         MaxChar         =   2
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         NumRighe        =   0
         MaxWidth        =   3
         CanRequired     =   0   'False
      End
      Begin PRJFW_CHECKBOX.TMS_CHECKBOX CHK_PARTBIS 
         Height          =   300
         Left            =   1140
         TabIndex        =   7
         Top             =   2760
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         IsDbField       =   0   'False
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL10 
         Height          =   300
         Left            =   210
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2790
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Bis"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_NUMPART 
         Height          =   300
         Left            =   1170
         TabIndex        =   6
         Top             =   2400
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         Default         =   "100"
         IsDbField       =   0   'False
         MaxWidth        =   4
         MaxChar         =   5
         CanRequired     =   0   'False
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL5 
         Height          =   300
         Left            =   210
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2460
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "Num. partita"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL4 
         Height          =   300
         Left            =   210
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2130
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         Caption         =   "Sez. partita"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_ANNOPART 
         Height          =   300
         Left            =   1170
         TabIndex        =   4
         Top             =   1740
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         Default         =   "100"
         IsDbField       =   0   'False
         MaxWidth        =   4
         MaxChar         =   5
         CanRequired     =   0   'False
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL3 
         Height          =   300
         Left            =   210
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1800
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "Anno partita"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_CLIFOR 
         Height          =   300
         Left            =   1170
         TabIndex        =   3
         Top             =   1410
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         Default         =   "100"
         IsDbField       =   0   'False
         MaxWidth        =   4
         MaxChar         =   5
         CanRequired     =   0   'False
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL9 
         Height          =   300
         Left            =   210
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1470
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "CliFor"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL7 
         Height          =   300
         Left            =   1890
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1260
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   529
         Caption         =   "1 = fornitore"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL8 
         Height          =   300
         Left            =   1890
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1050
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   529
         Caption         =   "0 = cliente"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_TIPOCF 
         Height          =   300
         Left            =   1170
         TabIndex        =   2
         Top             =   1080
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         Default         =   "100"
         IsDbField       =   0   'False
         MaxWidth        =   4
         MaxChar         =   5
         CanRequired     =   0   'False
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL6 
         Height          =   300
         Left            =   210
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1140
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Tipo CF"
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00ECFFFF&
         Caption         =   "Dati relativi alla registrazione di pagamento / incasso"
         Height          =   435
         Left            =   3120
         TabIndex        =   24
         Top             =   240
         Width           =   2985
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00ECFFFF&
         Caption         =   "Dati relativi alla partita da pagare / incassare"
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2535
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   1170
         TabIndex        =   1
         Top             =   750
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
      Begin VB.Label Label1 
         Caption         =   "Ditta"
         Height          =   225
         Left            =   210
         TabIndex        =   22
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Valuta"
         Height          =   225
         Left            =   3150
         TabIndex        =   21
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Cambio"
         Height          =   225
         Left            =   3150
         TabIndex        =   20
         Top             =   1140
         Width           =   645
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_CAMBIO 
         Height          =   300
         Left            =   4320
         TabIndex        =   9
         Top             =   1080
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxWidth        =   7
         FormatMask      =   "#,##0.00"
         CanRequired     =   0   'False
      End
      Begin VB.Label Label15 
         Caption         =   "Num. reg."
         Height          =   225
         Left            =   3150
         TabIndex        =   19
         Top             =   2850
         Width           =   765
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREGPAG 
         Height          =   300
         Left            =   4320
         TabIndex        =   13
         Top             =   2790
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
      Begin PRJFW_EDITNUM.TxtEditNum TXT_RIGACONTPAG 
         Height          =   300
         Left            =   4320
         TabIndex        =   14
         Top             =   3120
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         Default         =   "100"
         IsDbField       =   0   'False
         MaxWidth        =   4
         MaxChar         =   5
         CanRequired     =   0   'False
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL1 
         Height          =   300
         Left            =   3150
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3150
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         Caption         =   "Riga cont."
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL2 
         Height          =   300
         Left            =   3150
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3480
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
         Caption         =   "Data reg."
      End
   End
End
Attribute VB_Name = "FRM_CHIUDISCADENZE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StrConnect           As Variant
Public CallingForm          As FRM_MAIN

Private Connessione         As ADODB.Connection
Private ClsGestEcPort       As EPBO_MEMPORTAF.CLSEP_MEMPORTAF
Private ClsGestEcPortInput  As EPBO_MEMPORTAF.CLSEP_INPUTMEMPORTAF

Private Gcls_Connect        As CLSFW_SetConnect

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

Private Sub Form_Load()
    On Error GoTo Err_Form_Load
    
    '
    ' Creo la connessione
    '
    Set Gcls_Connect = New CLSFW_SetConnect
    Set Connessione = New ADODB.Connection
    Set Connessione = Gcls_Connect.Gpr_GetConnect
    With Connessione
        .ConnectionString = StrConnect
        .Open
    End With
    
    '
    ' Istanzio il b.o. per la gestione dell'EC/Port
    '
    Set ClsGestEcPort = New EPBO_MEMPORTAF.CLSEP_MEMPORTAF
    Set ClsGestEcPortInput = New EPBO_MEMPORTAF.CLSEP_INPUTMEMPORTAF
    Set ClsGestEcPort.CPInput = ClsGestEcPortInput
    
    Set ClsGestEcPort.CPInput.Connessione = Connessione
    
Exit Sub
Err_Form_Load:
    MsgBox Err.Number & " - " & Err.Description, , "Form_Load"
    Exit Sub
End Sub

Private Sub TXT_CAUSALEPAG_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    Dim Pcls_Lookup As COBO_LOOKUPDECODE.CLSCO_LOOKUP
    
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Lookup = New COBO_LOOKUPDECODE.CLSCO_LOOKUP
    
    Pcls_Lookup.CausaliContabili
    
    str_SQL = Pcls_Lookup.StringaSQL
    Arr_Fields = Pcls_Lookup.ArrayFields
    Str_Caption = Pcls_Lookup.Titolo
    Str_Connect = StrConnect
    
    Set Pcls_Lookup = Nothing
    Err.Clear
End Sub

Private Sub TXT_VALUTA_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    Dim Pcls_Lookup As COBO_LOOKUPDECODE.CLSCO_LOOKUP
    
    On Error Resume Next
    
    Cancel = False
    
    Set Pcls_Lookup = New COBO_LOOKUPDECODE.CLSCO_LOOKUP
    
    Pcls_Lookup.Valute
    
    str_SQL = Pcls_Lookup.StringaSQL
    Arr_Fields = Pcls_Lookup.ArrayFields
    Str_Caption = Pcls_Lookup.Titolo
    Str_Connect = StrConnect
    
    Set Pcls_Lookup = Nothing
    Err.Clear
End Sub

Private Sub BUT_ESEGUI_Click()
    On Error GoTo Err_BUT_ESEGUI_Click
    
    ClsGestEcPort.CPInput.Ditta = TXT_DITTA.Text
    ClsGestEcPort.CPInput.TipoCF = TXT_TIPOCF.Text
    ClsGestEcPort.CPInput.CliFor = TXT_CLIFOR.Text
    ClsGestEcPort.CPInput.AnnoPartita = TXT_ANNOPART.Text
    ClsGestEcPort.CPInput.NumeroPartita = TXT_NUMPART.Text
    ClsGestEcPort.CPInput.SezionalePartita = TXT_SEZPART.Text
    ClsGestEcPort.CPInput.PartitaBis = CHK_PARTBIS.Text
    
    ClsGestEcPort.CPInput.ImportoPagato = NVL(TXT_PAGATO.Text, 0)
    
    ClsGestEcPort.CPInput.Valuta = TXT_VALUTA.Text
    ClsGestEcPort.CPInput.Cambio = TXT_CAMBIO.Text
    ClsGestEcPort.CPInput.NumRegPagamento = TXT_NUMREGPAG.Text
    ClsGestEcPort.CPInput.RigaContPagamento = TXT_RIGACONTPAG.Text
    ClsGestEcPort.CPInput.DataRegPagamento = TXT_DATAREGPAG.Text
    ClsGestEcPort.CPInput.CausalePagamento = TXT_CAUSALEPAG.Text
    ClsGestEcPort.CPInput.DescCauPagamento = TXT_DESCRCAUSALE.Text
    
    ClsGestEcPort.ChiudiScadenze
    
    If ClsGestEcPort.Stato <> 0 Then
        MsgBox "Errore in ClsGestEcPort.ChiudiScadenzaPerChiave: " & ClsGestEcPort.Errore
        Exit Sub
    End If
    
    MsgBox "Elaborazione terminata"
    
Exit Sub
Err_BUT_ESEGUI_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_ESEGUI_Click"
    Exit Sub
End Sub
