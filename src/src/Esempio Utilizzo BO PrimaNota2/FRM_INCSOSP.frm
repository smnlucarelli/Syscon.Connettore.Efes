VERSION 5.00
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.9#0"; "TMS_EDIT.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.9#0"; "TMS_EDITNUM.ocx"
Object = "{0EF4EA13-2617-11D2-A1C0-0060082875F9}#7.8#0"; "TMS_EDITDATE.ocx"
Begin VB.Form FRM_INCSOSP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inserimento incasso fattura sospesa"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "FRM_INCSOSP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMD_INSERT 
      Caption         =   "Esegui"
      Height          =   525
      Left            =   7950
      Picture         =   "FRM_INCSOSP.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1470
      Width           =   1305
   End
   Begin VB.Frame Frame4 
      Caption         =   "Avanzamento:"
      Height          =   1695
      Left            =   60
      TabIndex        =   10
      Top             =   2070
      Width           =   7635
      Begin VB.ListBox LST_AVANZAMENTO 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   90
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dati relativi alla testata:"
      Height          =   1335
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   9195
      Begin VB.Label Label3 
         Caption         =   "Num. reg. incasso"
         Height          =   225
         Left            =   4560
         TabIndex        =   13
         Top             =   690
         Width           =   1335
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG_INCASSO 
         Height          =   300
         Left            =   5910
         TabIndex        =   12
         Top             =   630
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
      Begin PRJFW_EDITDATE.TxtEditDate TXT_DATAREG 
         Height          =   300
         Left            =   5910
         TabIndex        =   1
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         IsCalendario    =   0   'False
         IsDbField       =   0   'False
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_TOTALEINC 
         Height          =   300
         Left            =   2130
         TabIndex        =   3
         Top             =   960
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   529
         IsDbField       =   0   'False
         MaxChar         =   13
         CanRequired     =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG_FATTURA 
         Height          =   300
         Left            =   2130
         TabIndex        =   2
         Top             =   630
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
         Left            =   2130
         TabIndex        =   0
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
      Begin VB.Label Label1 
         Caption         =   "Codice ditta"
         Height          =   225
         Left            =   90
         TabIndex        =   9
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Data reg."
         Height          =   225
         Left            =   4560
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Num. reg. ft. da incassare"
         Height          =   225
         Left            =   90
         TabIndex        =   7
         Top             =   690
         Width           =   1845
      End
      Begin VB.Label Label6 
         Caption         =   "Totale incassato"
         Height          =   225
         Left            =   90
         TabIndex        =   6
         Top             =   990
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FRM_INCSOSP"
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

Private Pobj_PNota                      As CLSCG_PNOTACHECK

Public StrConnect                       As Variant
Private Connessione                     As ADODB.Connection
Public CallingForm                      As FRM_MAIN

Private Sub Form_Activate()
    On Error Resume Next
    
    TXT_DITTA.Text = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
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

'
' Ritorna il primo numero progressivo disponibile per il pagamento
'
Public Function GetNumPag(Ditta As Variant, NumRegFattura As Variant) As Variant
    Dim PRecSet         As ADODB.Recordset
    Dim StrSql          As Variant
    Dim UltimoNumPag    As Variant
    
    On Error Resume Next
    
    StrSql = "SELECT MAX(CG49_NUMPAG) AS ULTIMONUMPAG" & _
            " FROM CG49_PAGSOSP WITH (NOLOCK)" & _
            " WHERE CG49_DITTA_CG18 = " & Ditta & _
            " AND CG49_NUMREG_CG41 = '" & NumRegFattura & "'"
    
    Set PRecSet = Connessione.Execute(StrSql)
    
    UltimoNumPag = NVL(PRecSet.Fields("ULTIMONUMPAG").Value, 0)
    
    GetNumPag = UltimoNumPag + 1
    
    PRecSet.Close
    Set PRecSet = Nothing
    Err.Clear
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

Private Sub CalcolaImposta(ImponibileLordo As Variant, _
                           CodiceAliquota As Variant, _
                           ByRef ImponibileNetto As Variant, _
                           ByRef Imposta As Variant)
    
    On Error GoTo Err_CalcolaImposta
    
    If Pobj_PNota Is Nothing Then
        Set Pobj_PNota = New CLSCG_PNOTACHECK
    End If
    
    Pobj_PNota.CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    Pobj_PNota.Valuta = "EURO"
    Pobj_PNota.SommaImponibiliPercIva = ImponibileLordo
    Pobj_PNota.SommaImpostePercIva = 0
    Pobj_PNota.CodiceAliquota = CodiceAliquota
    Pobj_PNota.IndicatoreTipoRegistro = Vendite
    Pobj_PNota.TipoCalcoloImposta = Scorporo
    Pobj_PNota.IndicatoreProRata = 0 ' Non gestita
    Pobj_PNota.IndicatoreDetrIva = 2 ' Distinta dal costo
    
    Set Pobj_PNota.GConnect = Connessione
    Pobj_PNota.Sconnect = StrConnect
    
    Pobj_PNota.CalcolaImposta
    ImponibileNetto = ImponibileLordo - NVL(Pobj_PNota.Imposta, 0)
    Imposta = NVL(Pobj_PNota.Imposta, 0)
    
Exit Sub
Err_CalcolaImposta:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub CMD_INSERT_Click()
    Dim StatoFinale             As StatoFinaleEnum
    Dim SQL_CG48                As Variant
    Dim RecSet_CG48             As ADODB.Recordset
    Dim TotaleDaAssegnare       As Variant
    Dim Residuo                 As Variant
    Dim ImponibileNetto         As Variant
    Dim Imposta                 As Variant
    Dim ImponibileDaIncassare   As Variant
    Dim ImpostaDaIncassare      As Variant
    Dim Sql_CG41_Fattura        As Variant
    Dim RecSet_CG41_Fattura     As ADODB.Recordset
    Dim Perc                    As Variant
    Dim ImponibileLordoDaIncassare  As Variant
    Dim TotaleIVA                   As Variant
    
    On Error GoTo Err_CMD_INSERT_Click
    
    '
    ' Pulisco il list box che segnala le operazioni
    '
    LST_AVANZAMENTO.Clear
    LST_AVANZAMENTO.Refresh
    
    '
    ' Recupero i dati della testata della fattura sospesa
    '
    Sql_CG41_Fattura = "SELECT *" & _
                      " FROM CG41_PRIMANOTA WITH (NOLOCK)" & _
                      " WHERE CG41_DITTA_CG18 = " & TXT_DITTA.Text & _
                      " AND CG41_NUMREG = '" & TXT_NUMREG_FATTURA.Text & "'"
    
    Set RecSet_CG41_Fattura = Connessione.Execute(Sql_CG41_Fattura, , adCmdText)
    
    If RecSet_CG41_Fattura.RecordCount = 0 Then
        MsgBox "Non esiste la fattura da incassare."
        Exit Sub
    End If
    
    '
    ' Valorizzo le proprietà della classe che gestisce la prima nota
    '
    Pcls_PrimaNota.Status = tsInsert
    Set Pcls_PrimaNota.ActiveInterface = CallingForm.ActiveInterface
    Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
    Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione
    
    '
    ' Inserimento testata incasso
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata
        .ProvenienzaRegistrazione = DocumentoPrimaNota
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegistrazione = Null
        .DataCreazioneVariazione = Now
        .DataRegistrazione = TXT_DATAREG.Text
        
        .CodiceSezionale = RecSet_CG41_Fattura.Fields("CG41_SEZIONALE").Value
        .NumeroDocumento = 0
        
        .ContoCliFor = Null
        .DataRegIva = TXT_DATAREG.Text
        
        .DataDocumentoOrigine = RecSet_CG41_Fattura.Fields("CG41_DATADOC").Value
        .CodiceCausale = "102"
        .DescrCausale = "Test incasso fattura sospesa"
        .NumeroDocumentoOrigine = RecSet_CG41_Fattura.Fields("CG41_NUMDOCORIG").Value
        
        .CodiceValuta = "EURO"
        
        .IndicatoreTipoRegistro = EnumTipoRegistro.Vendite
        .FlagSegnoIva = SegnoIvaEnum.Positivo
        
        .ImportoDocumento = TXT_TOTALEINC.Text
        .ImportoValuta = 0
        
        .IndicatoreProvRegistrazione = PrimaNota
        .IndicatoreTipoMovimento = Consolidato
        .TipoDocumento = Null
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRegistrazione
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRegistrazione"
        Exit Sub
    End If
    
    TXT_NUMREG_INCASSO.Text = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
    
    '
    ' Inserimento riga incasso (cliente in avere)
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegistrazione = TXT_NUMREG_INCASSO.Text
        .NumeroRigaCont = 1
        .Conto = Left(CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroClienti, 2) & Right(String(8, "0") & Trim(CStr(NVL(RecSet_CG41_Fattura.Fields("CG41_CLIFOR_CG44").Value, 0))), 8)
        .CodiceValuta = "EURO"
        .Segno = IndicatoreDareAvere.Avere
        .Importo = TXT_TOTALEINC.Text
        .Imponibile = .Importo
        .IMPORTOVAL = 0
        .ImponibileVal = 0
        .CodiceAliquota = Null
        .IndicatoreTipoOperazione = CGBO_TIPI.TipoOperazione.PagamentoDocumento_DareFornitoreAvereCliente
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
        Exit Sub
    End If
    
    '
    ' Inserimento testata incasso (dati per sospensione imposta)
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestataPagSosp
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegDocOrigine = TXT_NUMREG_FATTURA.Text
        .NumeroPagamento = GetNumPag(.CodiceDitta, .NumeroRegDocOrigine)
        .DataPagamento = TXT_DATAREG.Text
        .NumeroRegistrazione = TXT_NUMREG_INCASSO.Text
        .ImportoTotale = TXT_TOTALEINC.Text
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRegistrazionePagSosp
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRegistrazionePagSosp"
        Exit Sub
    End If
    
    '
    ' Determino la percentuale dell'incassato sul totale fattura
    '
    Perc = NVL(TXT_TOTALEINC.Text, 0) / NVL(RecSet_CG41_Fattura.Fields("CG41_IMPTOTALE").Value, 0)
    
    '
    ' Recupero le righe della fattura da incassare relative all'IVA sospesa
    '
    SQL_CG48 = "SELECT *" & _
              " FROM CG48_MOVSOSP WITH (NOLOCK)" & _
              "    INNER JOIN CG28_TABCODIVA WITH (NOLOCK)" & _
              "        ON CG28_CODICE = CG48_CODICE_CG28" & _
              " WHERE CG48_DITTA_CG18 = " & TXT_DITTA.Text & _
              " AND CG48_NUMREG_CG41 = '" & TXT_NUMREG_FATTURA.Text & "'" & _
              " ORDER BY CG28_PERCIVA DESC"
    
    Set RecSet_CG48 = Connessione.Execute(SQL_CG48, , adCmdText)
    
    '
    ' Ripartisco il totale incassato sulle righe (proporzionamento)
    '
    If RecSet_CG48.RecordCount > 0 Then
        TotaleIVA = 0
        
        While Not RecSet_CG48.EOF
            
            ImponibileLordoDaIncassare = (NVL(RecSet_CG48.Fields("CG48_IMPONIBILE").Value, 0) + _
                                          NVL(RecSet_CG48.Fields("CG48_IMPOSTA").Value, 0)) * Perc
            
            CalcolaImposta ImponibileLordoDaIncassare, _
                           RecSet_CG48.Fields("CG48_CODICE_CG28").Value, _
                           ImponibileNetto, _
                           Imposta
            
            ImponibileDaIncassare = ImponibileNetto
            ImpostaDaIncassare = Imposta
            
            TotaleIVA = TotaleIVA + NVL(Imposta, 0)
            
            With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglioPagSosp
                .CodiceDitta = TXT_DITTA.Text
                .NumeroRegDocOrigine = TXT_NUMREG_FATTURA.Text
                .ProgressivoRiga = RecSet_CG48.Fields("CG48_PROGRIGA").Value
                
                .Imponibile = ImponibileDaIncassare
                .Imposta = ImpostaDaIncassare
                .ImponibileVal = 0
                .ImpostaVal = 0
                
                .NumeroPagamento = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestataPagSosp.NumeroPagamento
                .CodiceAliquota = RecSet_CG48.Fields("CG48_CODICE_CG28").Value
                
                '
                ' La proprietà IndicatoreTipo deve essere valorizzata in questo modo:
                ' - EnumTipoRegistro.ExSospesaVendite (nel caso di incasso fattura di vendita sospesa)
                ' - EnumTipoRegistro.ExSospesaAcquisti (nel caso di pagamento fattura di acquisto sospesa)
                ' - EnumTipoRegistro.ExSospesaCorrispettivi (nel caso di incasso corrispettivo sospesa)
                '
                .IndicatoreTipo = EnumTipoRegistro.ExSospesaVendite
            End With
            
            Pcls_PrimaNota.PGestRegPN.InserisciRigaPagSosp
            If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
                MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRigaPagSosp"
                Exit Sub
            End If
            
            RecSet_CG48.MoveNext
        Wend
    End If
    
    '
    ' Inserimento riga IVA sospesa (dare)
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegistrazione = TXT_NUMREG_INCASSO.Text
        .NumeroRigaCont = 2
        .Conto = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoIvaInSospensione
        .CodiceValuta = "EURO"
        .Segno = IndicatoreDareAvere.Dare
        .Importo = TotaleIVA
        .Imponibile = .Importo
        .IMPORTOVAL = 0
        .ImponibileVal = 0
        .CodiceAliquota = Null
        .IndicatoreTipoOperazione = CGBO_TIPI.TipoOperazione.GirocontoIvaSospesa
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
        Exit Sub
    End If
    
    '
    ' Inserimento riga IVA vendite (avere)
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegistrazione = TXT_NUMREG_INCASSO.Text
        .NumeroRigaCont = 3
        .Conto = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoIvaVendite
        .CodiceValuta = "EURO"
        .Segno = IndicatoreDareAvere.Avere
        .Importo = TotaleIVA
        .Imponibile = .Importo
        .IMPORTOVAL = 0
        .ImponibileVal = 0
        .CodiceAliquota = Null
        .IndicatoreTipoOperazione = CGBO_TIPI.TipoOperazione.GirocontoIvaSospesa
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
        Exit Sub
    End If
    
'''    '
'''    ' "Spalmo" il totale incassato sulle righe
'''    '
'''    If RecSet_CG48.RecordCount > 0 Then
'''        TotaleDaAssegnare = NVL(TXT_TOTALEINC.Text, 0)
'''
'''        While TotaleDaAssegnare > 0 And Not RecSet_CG48.EOF
'''            If NVL(RecSet_CG48.Fields("CG48_IMPONIBILE").Value, 0) + _
'''               NVL(RecSet_CG48.Fields("CG48_IMPOSTA").Value, 0) <= TotaleDaAssegnare Then
'''                '
'''                ' La riga può essere incassata totalmente
'''                '
'''                ImponibileDaIncassare = NVL(RecSet_CG48.Fields("CG48_IMPONIBILE").Value, 0)
'''                ImpostaDaIncassare = NVL(RecSet_CG48.Fields("CG48_IMPOSTA").Value, 0)
'''            Else
'''                '
'''                ' La riga viene incassata parzialmente
'''                '
'''                Residuo = TotaleDaAssegnare
'''
'''                CalcolaImposta Residuo, _
'''                               RecSet_CG48.Fields("CG48_CODICE_CG28").Value, _
'''                               ImponibileNetto, _
'''                               Imposta
'''
'''                ImponibileDaIncassare = ImponibileNetto
'''                ImpostaDaIncassare = Imposta
'''            End If
'''
'''            TotaleDaAssegnare = TotaleDaAssegnare - ImponibileDaIncassare - ImpostaDaIncassare
'''
'''            With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglioPagSosp
'''                .CodiceDitta = TXT_DITTA.Text
'''                .NumeroRegDocOrigine = TXT_NUMREG_FATTURA.Text
'''                .ProgressivoRiga = RecSet_CG48.Fields("CG48_PROGRIGA").Value
'''
'''                .Imponibile = ImponibileDaIncassare
'''                .Imposta = ImpostaDaIncassare
'''                .ImponibileVal = 0
'''                .ImpostaVal = 0
'''
'''                .NumeroPagamento = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestataPagSosp.NumeroPagamento
'''                .CodiceAliquota = RecSet_CG48.Fields("CG48_CODICE_CG28").Value
'''
'''                '
'''                ' La proprietà IndicatoreTipo deve essere valorizzata in questo modo:
'''                ' - EnumTipoRegistro.ExSospesaVendite (nel caso di incasso fattura di vendita sospesa)
'''                ' - EnumTipoRegistro.ExSospesaAcquisti (nel caso di pagamento fattura di acquisto sospesa)
'''                ' - EnumTipoRegistro.ExSospesaCorrispettivi (nel caso di incasso corrispettivo sospesa)
'''                '
'''                .IndicatoreTipo = EnumTipoRegistro.ExSospesaVendite
'''            End With
'''
'''            Pcls_PrimaNota.PGestRegPN.InserisciRigaPagSosp
'''            If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
'''                MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRigaPagSosp"
'''                Exit Sub
'''            End If
'''
'''            RecSet_CG48.MoveNext
'''        Wend
'''    End If
    
    '
    ' Inserimento riga incasso (cassa in dare)
    '
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
        .CodiceDitta = TXT_DITTA.Text
        .NumeroRegistrazione = TXT_NUMREG_INCASSO.Text
        .NumeroRigaCont = 4
        .Conto = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoCassa
        .CodiceValuta = "EURO"
        .Segno = IndicatoreDareAvere.Dare
        .Importo = TXT_TOTALEINC.Text
        .Imponibile = .Importo
        .IMPORTOVAL = 0
        .ImponibileVal = 0
        .CodiceAliquota = Null
        .IndicatoreTipoOperazione = CGBO_TIPI.TipoOperazione.PagamentoDocumento_ImportoPagatoCassa
    End With
    
    Pcls_PrimaNota.PGestRegPN.InserisciRiga
    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
        Exit Sub
    End If
    
    '
    ' Registro in database
    '
    Pcls_PrimaNota.InserisciPrimaNota StatoFinale
    
    If Pcls_PrimaNota.Stato <> tsOK Then
        MsgBox Pcls_PrimaNota.Errore & " in Pcls_PrimaNota.InserisciPrimaNota"
        Exit Sub
    End If
    
    MsgBox "La registrazione è stata inserita"
    
Exit Sub
Err_CMD_INSERT_Click:
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
End Sub

