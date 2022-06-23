VERSION 5.00
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.9#0"; "TMS_EDIT.ocx"
Begin VB.Form FRM_RITACCPAG 
   Caption         =   "Gestione ritenute acconto - pagamento"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   Icon            =   "FRM_RITACCPAG.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4335
   ScaleMode       =   0  'User
   ScaleWidth      =   8709.663
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Dati relativi alla ritenuta:"
      Height          =   1065
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   5115
      Begin VB.CommandButton BUT_CHIAMA_BO 
         Caption         =   "-> BO RITENUTE"
         Height          =   345
         Left            =   3210
         Picture         =   "FRM_RITACCPAG.frx":27A2
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   1785
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   1440
         TabIndex        =   0
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
      Begin VB.Label Label1 
         Caption         =   "Ditta"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   735
      End
   End
End
Attribute VB_Name = "FRM_RITACCPAG"
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

Private Sub BUT_CHIAMA_BO_Click()
    Dim ClsPagRit                   As CGBO_PAGRITENUTA.CLSCG_PAGRITENUTA
'    Dim ClsMemPortaf                As EPBO_MEMPORTAF.CLSEP_MEMPORTAF
'    Dim ClsMemPortafInput           As EPBO_MEMPORTAF.CLSEP_INPUTMEMPORTAF
    Dim RecSetScadenze              As ADODB.Recordset
    Dim SqlScadenze                 As Variant
    Dim RecSetScadenze_Importi      As ADODB.Recordset
    Dim Indice                      As Variant
    
    On Error GoTo ErrTrap
    
    '
    ' Creo la connessione
    '
    Set Connessione = New ADODB.Connection
    Connessione.ConnectionString = StrConnect
    Connessione.CursorLocation = adUseClient
    Connessione.Open
    
    '
    ' Definisco il recordset delle scadenze da pagare
    '
    SqlScadenze = "SELECT *" & vbCrLf & _
                 " FROM EF01_SCADENZE WITH (NOLOCK)" & vbCrLf & _
                 " WHERE EF01_DITTA_CG18 = " & NVL(TXT_DITTA.Text, 0) & vbCrLf & _
                 " AND EF01_TIPOCF_CG44 = 1" & vbCrLf & _
                 " AND EF01_CLIFOR_CG44 = 5"
    Set RecSetScadenze = Connessione.Execute(SqlScadenze, , adCmdText)
    
    '
    ' Definisco la struttura del recordset con gli importi pagati
    '
    Set RecSetScadenze_Importi = New ADODB.Recordset
    
    With RecSetScadenze_Importi
        .Fields.Append "EF01_DITTA_CG18", adDecimal, 5, adFldIsNullable
        .Fields.Append "EF01_NUMREG_CO99", adChar, 12, adFldIsNullable
        
        .Fields.Append "EF01_RIGACONT_CG42", adDecimal, , adFldIsNullable
        .Fields("EF01_RIGACONT_CG42").Precision = 5
        .Fields("EF01_RIGACONT_CG42").NumericScale = 0
        
        .Fields.Append "EF01_PROGEC_EC01", adDecimal, 6, adFldIsNullable
        .Fields.Append "EF01_PROGEFF", adDecimal, 6, adFldIsNullable
        .Fields.Append "IMPORTOPAGATO", adDecimal, 20, adFldIsNullable
        .Fields("IMPORTOPAGATO").Precision = 13
        .Fields("IMPORTOPAGATO").NumericScale = 2
        .Fields.Append "ABBUONO", adDecimal, 20, adFldIsNullable
        .Fields("ABBUONO").Precision = 13
        .Fields("ABBUONO").NumericScale = 2
        .Fields.Append "ACCONTO", adDecimal, 20, adFldIsNullable
        .Fields("ACCONTO").Precision = 13
        .Fields("ACCONTO").NumericScale = 2
        .Fields.Append "NUMSCAD", adDecimal, 20, adFldIsNullable
        .Fields.Append "RITACCONTO", adDecimal, 20, adFldIsNullable
        .Fields("RITACCONTO").Precision = 13
        .Fields("RITACCONTO").NumericScale = 2
        .Fields.Append "EF01_IMPEFF_S", adDecimal, 20, adFldIsNullable
        .Fields("EF01_IMPEFF_S").Precision = 13
        .Fields("EF01_IMPEFF_S").NumericScale = 2
        .Fields.Append "EF01_IMPEFFVAL_S", adDecimal, 20, adFldIsNullable
        .Fields("EF01_IMPEFFVAL_S").Precision = 13
        .Fields("EF01_IMPEFFVAL_S").NumericScale = 2
        .Fields.Append "EF01_IMPRITACC", adDecimal, 20, adFldIsNullable
        .Fields("EF01_IMPRITACC").Precision = 13
        .Fields("EF01_IMPRITACC").NumericScale = 2
        .Fields.Append "EF01_IMPRITACCVAL", adDecimal, 20, adFldIsNullable
        .Fields("EF01_IMPRITACCVAL").Precision = 13
        .Fields("EF01_IMPRITACCVAL").NumericScale = 2
        .Fields.Append "EF01_NUMRATA", adDecimal, 3, adFldIsNullable
        .Fields.Append "EF01_TIPOEFF", adDecimal, 2, adFldIsNullable
        .Fields.Append "CAUSALEAGGIUNTIVA", adVariant, 240, adFldIsNullable
        
        .Fields.Append "ABBUONOSUACCONTO", adDecimal, , adFldIsNullable
        .Fields("ABBUONOSUACCONTO").Precision = 13
        .Fields("ABBUONOSUACCONTO").NumericScale = 2
        
        .CursorType = adOpenKeyset
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .Open
    End With
    
    '
    ' Popolo il recordset relativo agli importi pagati
    '
    If RecSetScadenze.RecordCount = 0 Then
        MsgBox "Non ci sono scadenze da chiudere"
        Exit Sub
    End If
    
    Indice = 0
    RecSetScadenze.MoveFirst
    While Not RecSetScadenze.EOF
        
        With RecSetScadenze_Importi
            
            Indice = Indice + 1
            
            .AddNew
            
            .Fields("EF01_DITTA_CG18").Value = RecSetScadenze.Fields("EF01_DITTA_CG18").Value
            .Fields("EF01_NUMREG_CO99").Value = RecSetScadenze.Fields("EF01_NUMREG_CO99").Value
            .Fields("EF01_RIGACONT_CG42").Value = RecSetScadenze.Fields("EF01_RIGACONT_CG42").Value
            .Fields("EF01_PROGEC_EC01").Value = RecSetScadenze.Fields("EF01_PROGEC_EC01").Value
            .Fields("EF01_PROGEFF").Value = RecSetScadenze.Fields("EF01_PROGEFF").Value
            
            .Fields("ABBUONO").Value = 0
            .Fields("ACCONTO").Value = 0
            .Fields("NUMSCAD").Value = Indice
            .Fields("RITACCONTO").Value = RecSetScadenze.Fields("EF01_IMPRITACC").Value
            
            .Fields("EF01_IMPEFF_S").Value = RecSetScadenze.Fields("EF01_IMPEFF_S").Value
            .Fields("EF01_IMPEFFVAL_S").Value = RecSetScadenze.Fields("EF01_IMPEFFVAL_S").Value
            .Fields("EF01_IMPRITACC").Value = RecSetScadenze.Fields("EF01_IMPRITACC").Value
            .Fields("EF01_IMPRITACCVAL").Value = RecSetScadenze.Fields("EF01_IMPRITACCVAL").Value
            .Fields("EF01_NUMRATA").Value = RecSetScadenze.Fields("EF01_NUMRATA").Value
            .Fields("EF01_TIPOEFF").Value = RecSetScadenze.Fields("EF01_TIPOEFF").Value
            
            .Fields("IMPORTOPAGATO").Value = RecSetScadenze.Fields("EF01_IMPEFFNETPREV_S").Value
            
            .Update
        End With
        
        RecSetScadenze.MoveNext
    Wend
    
    '
    ' Creo l'istanza del b.o. mem. portaf.
    '
'    Set ClsMemPortaf = New EPBO_MEMPORTAF.CLSEP_MEMPORTAF
'    Set ClsMemPortafInput = New EPBO_MEMPORTAF.CLSEP_INPUTMEMPORTAF
'    Set ClsMemPortaf.CPInput = ClsMemPortafInput
'    Set ClsMemPortaf.CPInput.Connessione = Connessione
'    ClsMemPortaf.CPInput.StrConnessione = StrConnect
'    ClsMemPortaf.CPInput.Ditta = NVL(TXT_DITTA.Text, 0)
'    ClsMemPortaf.CPInput.TipoCF = 1
'    ClsMemPortaf.CPInput.CliFor = 5
'
'    ClsMemPortaf.CreaRecordsetImportiPagatiPerPagamentoEffettiProfessionisti RecSetScadenze
    
    '
    ' Creo l'istanza del b.o. pag. ritenute
    '
    Set ClsPagRit = New CGBO_PAGRITENUTA.CLSCG_PAGRITENUTA
    
    '
    ' Definizione proprietà per b.o. pag. ritenute
    '
    ClsPagRit.CPInput.StrConnessione = Connessione.ConnectionString
    Set ClsPagRit.CPInput.Connessione = Connessione
    
    ClsPagRit.ClsGestRegPN.CPInput.DatiTestata.CodiceDitta = TXT_DITTA.Text
    ClsPagRit.ClsGestRegPN.CPInput.DatiTestata.CodiceValuta = "EURO"
    ClsPagRit.ClsGestRegPN.CPInput.DatiTestata.Cambio = 1
    ClsPagRit.ClsGestRegPN.CPInput.DatiTestata.DataRegistrazione = CDate("14/10/2011")
    ClsPagRit.ClsGestRegPN.CPInput.DatiTestata.CodiceCausale = "101"
    ClsPagRit.CPInput.ContoPagamento = "1600010100"
    
    Set ClsPagRit.CPInput.RecSetEffettiDaPagare = RecSetScadenze_Importi
    
    ClsPagRit.DefinisciRegistrazionePagamentoEffettiProfessionisti
    
    If ClsPagRit.Stato <> tsOK Then
        MsgBox "Errore in DefinisciRegistrazionePagamentoEffettiProfessionisti:" & vbCrLf & _
               "Stato: " & ClsPagRit.Stato & vbCrLf & _
               "Errore: " & ClsPagRit.Errore
        Exit Sub
    End If
    
    
    


Exit Sub
ErrTrap:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_CHIAMA_BO_Click"
    Exit Sub
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    TXT_DITTA.Text = 1
    
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Form_Load
    
    Set ClsRitenute = New CGBO_RITENUTE.CLSCG_GESTRITENUTE
    ClsRitenute.CPInput.Sconnect = StrConnect
    
Exit Sub
Err_Form_Load:
    MsgBox Err.Number & " - " & Err.Description, , "Form_Load"
    Exit Sub
End Sub
