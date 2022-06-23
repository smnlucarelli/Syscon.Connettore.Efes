VERSION 5.00
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.11#0"; "TMS_EDITM.ocx"
Begin VB.Form FRM_INCASSO_SOSP 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMD_ELABORA 
      Caption         =   "ELABORA"
      Height          =   615
      Left            =   4470
      TabIndex        =   2
      Top             =   90
      Width           =   1965
   End
   Begin PRJFW_EDITM.TXT_EDITM TXT_CLIENTE 
      Height          =   300
      Left            =   1950
      TabIndex        =   1
      Top             =   180
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      IsLookup        =   -1  'True
      DisplayFormat   =   "Maiuscolo"
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      NumRighe        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "CLIENTE"
      Height          =   255
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1605
   End
End
Attribute VB_Name = "FRM_INCASSO_SOSP"
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

Private ClsIncassoScad                  As CGUO_INCASSOSCAD.CLSCG_INCASSOSCAD

Private DictNumReg                      As New Dictionary

Private Sub CMD_ELABORA_Click()
    Dim Sql                         As Variant
    Dim RecSet                      As ADODB.Recordset
    Dim RecSetClone                 As ADODB.Recordset
    Dim TotIncTestataPN             As Variant
    Dim TotIncScadenze              As Variant
    Dim NumRegInc                   As Variant
    Dim RecSetScadenzePagInc        As ADODB.Recordset
    Dim ProgEffCorrente             As Variant
    Dim Campo                       As ADODB.Field
    
    Dim RecSetFattureDaIncassare    As ADODB.Recordset
    Dim RecSetScadenzeDaIncassare   As ADODB.Recordset
    Dim RecSetContropartite         As ADODB.Recordset
    
    Dim RecSetContropartitePagInc   As ADODB.Recordset
    
    Dim OldImponibileIncassato      As Variant
    Dim OldImpostaIncassata         As Variant
    
    Dim NumRigaCont                 As Variant
    
    Dim IvaSosp                     As Variant
    Dim NumRigaContCliente          As Variant
    
    Dim CG41_DATAREGIVA             As Variant
    Dim CG41_DATAREG                As Variant
    Dim CG41_SEZIONALE              As Variant

    Dim IvaVendite                  As Variant
    Dim IvaVenditeSosp              As Variant

    Dim TipoEsigIVA                 As Variant
    Dim PercVent                    As Variant

    Sql = "SELECT *" & _
         " FROM EF01_SCADENZE" & _
         " WHERE EF01_DITTA_CG18 = " & CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta & _
         " AND EF01_TIPOCF_CG44 = 0" & _
         " AND EF01_CLIFOR_CG44 = " & TXT_CLIENTE.Text & _
         " AND EF01_INDSTATO_S = 0" & _
         " ORDER BY EF01_DITTA_CG18, EF01_NUMREG_CO99, EF01_PROGEFF"
    
    Set RecSet = Connessione.Execute(Sql, , adCmdText)
    
    '
    ' Determino il totale da incassare
    '
    TotIncTestataPN = 0
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        While Not RecSet.EOF
            TotIncTestataPN = TotIncTestataPN + RecSet.Fields("EF01_IMPEFFNETPREV_S").Value
            RecSet.MoveNext
        Wend
        RecSet.MoveFirst
    End If
    
    '
    ' Inserisco la testata del movimento
    '
    
    '
    ' Valorizzo le proprietà della classe che gestisce la prima nota
    '
    Pcls_PrimaNota.Status = tsInsert
    Set Pcls_PrimaNota.ActiveInterface = CallingForm.ActiveInterface
    Pcls_PrimaNota.PGestRegPN.CPInput.Sconnect = StrConnect
    Set Pcls_PrimaNota.PGestRegPN.CPInput.GConnect = Connessione
    
    With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata
        .ProvenienzaRegistrazione = DocumentoPrimaNota
        .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
        .NumeroRegistrazione = Null
        .DataCreazioneVariazione = Now
        .DataRegistrazione = "26/01/2011"
        
        .CodiceSezionale = "00"
        .NumeroDocumento = 0
        
        .ContoCliFor = Null
        .DataRegIva = .DataRegistrazione
        
        .DataDocumentoOrigine = Null
        .CodiceCausale = "102"
        .DescrCausale = "Test incasso fattura sospesa"
        .NumeroDocumentoOrigine = "doc orig"
        
        .CodiceValuta = "EURO"
        
        .IndicatoreTipoRegistro = EnumTipoRegistro.Vendite
        .FlagSegnoIva = SegnoIvaEnum.Positivo
        
        .ImportoDocumento = TotIncTestataPN
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
    
    NumRegInc = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestata.NumeroRegistrazione
    NumRigaCont = 0
    
    If Not RecSet.EOF Then
        RecSet.MoveFirst
        While Not RecSet.EOF
            '
            ' Inserimento riga cliente in avere (in dare se nota credito)
            '
            NumRigaCont = NumRigaCont + 1
            With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
                .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
                .NumeroRegistrazione = NumRegInc
                .NumeroRigaCont = NumRigaCont
                .Conto = Left(CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroClienti, 2) & Right(String(8, "0") & Trim(CStr(NVL(RecSet.Fields("EF01_CLIFOR_CG44").Value, 0))), 8)
                .CodiceValuta = "EURO"
                .Segno = IndicatoreDareAvere.Avere
                .Importo = RecSet.Fields("EF01_IMPEFFNETPREV_S").Value
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
            
            NumRigaContCliente = NumRigaCont
            
            '
            ' Elaborazione per scadenze ad esigibilità IVA differita
            '
            Set ClsIncassoScad = New CGUO_INCASSOSCAD.CLSCG_INCASSOSCAD
            
            If ClsIncassoScad.CreaStrutturaRecordsetScadenze(RecSetScadenzePagInc, Connessione) = True Then
                '
                ' Determino il progressivo effetto corrente
                '
                ProgEffCorrente = RecSet.Fields("EF01_PROGEFF").Value
                
                '
                ' Definisco il recordset clone delle scadenze, filtrato per numreg
                '
                Set RecSetClone = RecSet.Clone
                RecSetClone.Filter = "EF01_NUMREG_CO99 = '" & RecSet.Fields("EF01_NUMREG_CO99").Value & "'"
                RecSetClone.Sort = "EF01_PROGEFF"
                
                If Not RecSetClone.EOF Then
                    RecSetClone.MoveFirst
                    
                    TotIncScadenze = 0
                    While Not RecSetClone.EOF
                        If RecSetClone.Fields("EF01_PROGEFF").Value <= ProgEffCorrente Then
                            RecSetScadenzePagInc.AddNew
                            
                            For Each Campo In RecSetClone.Fields
                                RecSetScadenzePagInc.Fields(Campo.Name).Value = Campo.Value
                            Next
                            
                            TotIncScadenze = TotIncScadenze + _
                                                NVL(RecSetClone.Fields("EF01_IMPEFF_S").Value, 0) + _
                                                NVL(RecSetClone.Fields("EF01_IMPRITACC").Value, 0)
                            
                            RecSetScadenzePagInc.Fields("IMPORTOPAGATO").Value = _
                                                NVL(RecSetClone.Fields("EF01_IMPEFF_S").Value, 0) + _
                                                NVL(RecSetClone.Fields("EF01_IMPRITACC").Value, 0)
                            
                            RecSetScadenzePagInc.Fields("ABBUONO").Value = 0
                        End If
                        
                        RecSetClone.MoveNext
                    Wend
                End If
                
                With ClsIncassoScad
                    .VisualizzaFormContropartite = False
                    .TipoCF = 0 ' cliente
                    .CliFor = TXT_CLIENTE.Text
                    .ModoApertura = tsInsert
                    .TotaleDaIncassare = TotIncScadenze ' totale incassato
                    .NumRegPagInc = NumRegInc ' numreg pagamento
                    .TipoElaborazione = ElaborazioneSospensioneImposta
                    .CodiceCausale = "102"
                    .CodiceValuta = "EURO"
                    .DataRegistrazione = "26/01/2011"
                    
                    Set .RecSetScadenzePagInc = RecSetScadenzePagInc
                    
                    .ElaboraScadenze Connessione, _
                                     RecSetFattureDaIncassare, _
                                     RecSetScadenzeDaIncassare, _
                                     RecSetContropartite
                    
                    If Not .AnnullataDefinizioneDatiIncasso Then
                        If Not RecSetContropartite Is Nothing Then
                            If RecSetContropartite.State = adStateOpen Then
                                If RecSetContropartite.RecordCount > 0 Then
                                    RecSetContropartite.MoveFirst
                                    
                                    While Not RecSetContropartite.EOF
                                        RecSetContropartite.Fields("CG49_RIGACONTPN_CG42").Value = NumRigaContCliente ' riga contabile cliente
                                        RecSetContropartite.MoveNext
                                    Wend
                                    
                                    RecSetContropartite.MoveFirst
                                End If
                            End If
                        End If
                    End If
                End With
                
                If RecSetContropartitePagInc Is Nothing Then
                    If Not RecSetContropartite Is Nothing Then
                        '
                        ' Sto elaborando la prima scadenza, quindi creo il recordset RecSetContropartitePagInc
                        ' con la stessa struttura e gli stessi dati di RecSetContropartite
                        '
                        Set RecSetContropartitePagInc = New ADODB.Recordset
                        For Each Campo In RecSetContropartite.Fields
                            RecSetContropartitePagInc.Fields.Append Campo.Name, _
                                                                    Campo.Type, _
                                                                    Campo.DefinedSize, _
                                                                    Campo.Attributes
                            
                            If Campo.Type = adDecimal Or Campo.Type = adNumeric Then
                                RecSetContropartitePagInc.Fields(Campo.Name).NumericScale = Campo.NumericScale
                                RecSetContropartitePagInc.Fields(Campo.Name).Precision = Campo.Precision
                            End If
                        Next
                        RecSetContropartitePagInc.Open
                        
                        If RecSetContropartite.State = adStateOpen Then
                            If RecSetContropartite.RecordCount > 0 Then
                                RecSetContropartite.MoveFirst
                                While Not RecSetContropartite.EOF
                                    RecSetContropartitePagInc.AddNew
                                    For Each Campo In RecSetContropartite.Fields
                                        RecSetContropartitePagInc.Fields(Campo.Name).Value = _
                                                       RecSetContropartite.Fields(Campo.Name).Value
                                    Next
                                    
                                    RecSetContropartite.MoveNext
                                Wend
                            End If
                        End If
                    End If
                Else
                    '
                    ' Storno dal recordset RecSetContropartite determinato da CGUO_INCASSOSCAD
                    ' gli importi pagati/incassati presenti in RecSetContropartitePagInc
                    ' (relativo alle scadenze precedenti già chiuse).
                    ' Contemporaneamente, aggiorno il recordset RecSetContropartitePagInc
                    ' con gli importi presenti in RecSetContropartite
                    '
                    If Not RecSetContropartite Is Nothing Then
                        If RecSetContropartite.State = adStateOpen Then
                            If RecSetContropartite.RecordCount > 0 Then
                                RecSetContropartite.MoveFirst
                                While Not RecSetContropartite.EOF
                                    '
                                    ' Sincronizzo il recordset delle contropartite pagate / incassate
                                    '
                                    RecSetContropartitePagInc.Filter = "CG48_DITTA_CG18 = " & RecSetContropartite.Fields("CG48_DITTA_CG18").Value & _
                                                                  " AND CG48_NUMREG_CG41 = '" & RecSetContropartite.Fields("CG48_NUMREG_CG41").Value & "'" & _
                                                                  " AND CG48_PROGRIGA = " & RecSetContropartite.Fields("CG48_PROGRIGA").Value
                                    
                                    If RecSetContropartitePagInc.EOF Then
                                        RecSetContropartitePagInc.Filter = adFilterNone
                                        
                                        RecSetContropartitePagInc.AddNew
                                        For Each Campo In RecSetContropartite.Fields
                                            RecSetContropartitePagInc.Fields(Campo.Name).Value = _
                                                           RecSetContropartite.Fields(Campo.Name).Value
                                        Next
                                    Else
                                        OldImponibileIncassato = NVL(RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value, 0)
                                        OldImpostaIncassata = NVL(RecSetContropartite.Fields("IMPOSTAINCASSATA").Value, 0)
                                        
                                        RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value = NVL(RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value, 0) - _
                                                                                                  NVL(RecSetContropartitePagInc.Fields("IMPONIBILEINCASSATO").Value, 0)
                                        RecSetContropartite.Fields("IMPOSTAINCASSATA").Value = NVL(RecSetContropartite.Fields("IMPOSTAINCASSATA").Value, 0) - _
                                                                                               NVL(RecSetContropartitePagInc.Fields("IMPOSTAINCASSATA").Value, 0)
                                        
                                        RecSetContropartitePagInc.Fields("IMPONIBILEINCASSATO").Value = OldImponibileIncassato
                                        RecSetContropartitePagInc.Fields("IMPOSTAINCASSATA").Value = OldImpostaIncassata
                                    End If
                                    
                                    RecSetContropartite.MoveNext
                                Wend
                                
                                RecSetContropartite.MoveFirst
                            End If
                        End If
                    End If
                End If
            End If
            
            '
            ' Determino il totale dell'iva da girocontare
            '
            IvaSosp = 0
            If Not RecSetContropartite Is Nothing Then
                If RecSetContropartite.State = adStateOpen Then
                    If RecSetContropartite.RecordCount > 0 Then
                        RecSetContropartite.MoveFirst
                        While Not RecSetContropartite.EOF
                            IvaSosp = IvaSosp + NVL(RecSetContropartite.Fields("IMPOSTAINCASSATA").Value, 0)
                            
                            RecSetContropartite.MoveNext
                        Wend
                    End If
                End If
            End If
            
            If IvaSosp <> 0 Then
                IvaVendite = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoIvaVendite
                IvaVenditeSosp = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoIvaInSospensione
                
                '
                ' Inserimento riga iva in sospensione in dare (in avere se nota credito)
                '
                NumRigaCont = NumRigaCont + 1
                With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
                    .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
                    .NumeroRegistrazione = NumRegInc
                    .NumeroRigaCont = NumRigaCont
                    .Conto = IvaVenditeSosp
                    .CodiceValuta = "EURO"
                    .Segno = IndicatoreDareAvere.Dare
                    .Importo = IvaSosp
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
                ' Inserimento riga iva vendite avere (in dare se nota credito)
                '
                NumRigaCont = NumRigaCont + 1
                With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
                    .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
                    .NumeroRegistrazione = NumRegInc
                    .NumeroRigaCont = NumRigaCont
                    .Conto = IvaVenditeSosp
                    .CodiceValuta = "EURO"
                    .Segno = IndicatoreDareAvere.Avere
                    .Importo = IvaSosp
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
                ' Inserimento della testata dell'incasso/pagamento esigibilità IVA differita
                '
                With Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestataPagSosp
                    .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
                    .NumeroRegDocOrigine = RecSet.Fields("EF01_NUMREG_CO99").Value
                    .NumeroPagamento = GetNumPag(.NumeroRegDocOrigine)
                    .DataPagamento = "26/01/2011"
                    .NumeroRegistrazione = NumRegInc
                    .RigaContPagamento = NumRigaContCliente
                    .ImportoTotale = CalcolaTotaleIncPagSosp(RecSetContropartite)
                    .ImportoTotaleVal = 0
                    .ImportoArrotondamento = 0
                    .ImportoArrotondamentoVal = 0
                End With
                
                Pcls_PrimaNota.PGestRegPN.InserisciRegistrazionePagSosp
                If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
                    MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in InserisciRegistrazionePagSosp"
                    Exit Sub
                End If
                
                If Not RecSetContropartite Is Nothing Then
                    If RecSetContropartite.State = adStateOpen Then
                        If RecSetContropartite.RecordCount > 0 Then
                            RecSetContropartite.MoveFirst
                            
                            While Not RecSetContropartite.EOF
                                With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglioPagSosp
                                    .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
                                    .NumeroRegDocOrigine = RecSet.Fields("EF01_NUMREG_CO99").Value
                                    .ProgressivoRiga = RecSetContropartite.Fields("CG48_PROGRIGA").Value
                                    
                                    If NVL(RecSetContropartite.Fields("CG41_FLGSEGNOIVA").Value, 0) = 0 Then
                                        '
                                        ' Fattura
                                        '
                                        If NVL(RecSet.Fields("EF01_CODICE_CG08").Value, "EURO") = "EURO" Then
                                            .Imponibile = RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value
                                            .Imposta = RecSetContropartite.Fields("IMPOSTAINCASSATA").Value
                                            .ImponibileVal = 0
                                            .ImpostaVal = 0
                                        Else
                                            .Imponibile = RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value
                                            .Imposta = RecSetContropartite.Fields("IMPOSTAINCASSATA").Value
                                            .ImponibileVal = RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value
                                            .ImpostaVal = RecSetContropartite.Fields("IMPOSTAINCASSATA").Value
                                        End If
                                    Else
                                        '
                                        ' Nota credito
                                        '
                                        If NVL(RecSet.Fields("EF01_CODICE_CG08").Value, "EURO") = "EURO" Then
                                            .Imponibile = -RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value
                                            .Imposta = -RecSetContropartite.Fields("IMPOSTAINCASSATA").Value
                                            .ImponibileVal = 0
                                            .ImpostaVal = 0
                                        Else
                                            .Imponibile = -RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value
                                            .Imposta = -RecSetContropartite.Fields("IMPOSTAINCASSATA").Value
                                            .ImponibileVal = -RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value
                                            .ImpostaVal = -RecSetContropartite.Fields("IMPOSTAINCASSATA").Value
                                        End If
                                    End If
                                    
                                    .NumeroPagamento = Pcls_PrimaNota.PGestRegPN.CPInput.DatiTestataPagSosp.NumeroPagamento
                                    
                                    .CodiceAliquota = RecSetContropartite.Fields("CG48_CODICE_CG28").Value
                                    
                                    '
                                    ' Determino i dati legati al codice IVA
                                    '
                                    TipoEsigIVA = 0
                                    PercVent = 0
                                    FindAliquota .CodiceAliquota, TipoEsigIVA, PercVent
                                    
                                    '
                                    ' Determino i dati della fattura origine
                                    '
                                    GetDatiFatturaOrigine RecSet.Fields("EF01_NUMREG_CO99").Value, CG41_DATAREGIVA, CG41_DATAREG, CG41_SEZIONALE
                                    
                                    Select Case NVL(RecSetContropartite.Fields("CG41_INDTIPOREG").Value, 0)
                                        Case EnumTipoRegistro.Acquisti
                                            Select Case TipoEsigIVA
                                                Case 0, 1
                                                    .IndicatoreTipo = EnumTipoRegistro.ExSospesaAcquisti
                                                Case 2
                                                    If NVL(CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiAttivita.IndVent(Year(NVL(CG41_DATAREGIVA, CG41_DATAREG)), , CG41_SEZIONALE), 0) <> 0 And _
                                                       NVL(PercVent, 0) <> 0 Then
                                                        '
                                                        ' Ventilazione
                                                        '
                                                        .IndicatoreTipo = EnumTipoRegistro.ExSospesaAcquistiMonteVentilazione_AltreSocieta
                                                    Else
                                                        .IndicatoreTipo = EnumTipoRegistro.ExSospesaAcquisti_AltreSocieta
                                                    End If
                                            End Select
                                        Case EnumTipoRegistro.Vendite
                                            Select Case TipoEsigIVA
                                                Case 0, 1
                                                    .IndicatoreTipo = EnumTipoRegistro.ExSospesaVendite
                                                Case 2
                                                    .IndicatoreTipo = EnumTipoRegistro.ExSospesaVendite_AltreSocieta
                                            End Select
                                        Case EnumTipoRegistro.Corrispettivi
                                            Select Case TipoEsigIVA
                                                Case 0, 1
                                                    .IndicatoreTipo = EnumTipoRegistro.ExSospesaCorrispettivi
                                                Case 2
                                                    .IndicatoreTipo = EnumTipoRegistro.ExSospesaCorrispettivi_AltreSocieta
                                            End Select
                                    End Select
                                End With
                                
                                If NVL(Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglioPagSosp.Imponibile, 0) <> 0 Or _
                                   NVL(Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglioPagSosp.Imposta, 0) <> 0 Or _
                                   NVL(Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglioPagSosp.ImponibileVal, 0) <> 0 Or _
                                   NVL(Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglioPagSosp.ImpostaVal, 0) <> 0 Then
                                    
                                    Pcls_PrimaNota.PGestRegPN.InserisciRigaPagSosp
                                    
                                    If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
                                        MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in InserisciRegistrazionePagSosp"
                                        Exit Sub
                                    End If
                                End If
                                
                                RecSetContropartite.MoveNext
                            Wend
                        End If
                    End If
                End If
            End If
            
            Set ClsIncassoScad = Nothing
            
            RecSet.MoveNext
        Wend
        
        '
        ' Inserimento riga banca in dare
        '
        NumRigaCont = NumRigaCont + 1
        With Pcls_PrimaNota.PGestRegPN.CPInput.DatiDettaglio
            .CodiceDitta = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
            .NumeroRegistrazione = NumRegInc
            .NumeroRigaCont = NumRigaCont
            .Conto = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.ContoCassa
            .CodiceValuta = "EURO"
            .Segno = IndicatoreDareAvere.Dare
            .Importo = TotIncTestataPN
            .IMPORTOVAL = 0
            .ImponibileVal = 0
            .CodiceAliquota = Null
            .IndicatoreTipoOperazione = CGBO_TIPI.PagamentoDocumento_ImportoPagatoCassa
        End With
        
        Pcls_PrimaNota.PGestRegPN.InserisciRiga
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.InserisciRiga"
            Exit Sub
        End If
        
        Pcls_PrimaNota.PGestRegPN.RegistraModifiche
        If Pcls_PrimaNota.PGestRegPN.Stato <> tsOK Then
            MsgBox Pcls_PrimaNota.PGestRegPN.Errore & " in Pcls_PrimaNota.PGestRegPN.RegistraModifiche"
            Exit Sub
        End If
        
        '
        ' Scrittura estratto conto / portafoglio
        '
        ScriviEcPort RecSet, NumRegInc
    
    End If
    
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
' Ritorna il primo numero progressivo disponibile per il pagamento
'
Private Function GetNumPag(NumRegFattura As Variant) As Variant
    Dim PRecSet         As ADODB.Recordset
    Dim StrSql          As Variant
    Dim UltimoNumPag    As Variant
    
    On Error Resume Next
    
    If DictNumReg.Exists(NumRegFattura) Then
        GetNumPag = DictNumReg(NumRegFattura) + 1
        DictNumReg(NumRegFattura) = GetNumPag
    Else
        StrSql = "SELECT MAX(CG49_NUMPAG) AS ULTIMONUMPAG" & _
                " FROM CG49_PAGSOSP WITH (NOLOCK)" & _
                " WHERE CG49_DITTA_CG18 = " & CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta & _
                " AND CG49_NUMREG_CG41 = '" & NumRegFattura & "'"
        
        Set PRecSet = Connessione.Execute(StrSql)
        
        UltimoNumPag = NVL(PRecSet.Fields("ULTIMONUMPAG").Value, 0)
        
        GetNumPag = UltimoNumPag + 1
        
        PRecSet.Close
        Set PRecSet = Nothing
        
        DictNumReg(NumRegFattura) = GetNumPag
    End If
    
    Err.Clear
End Function

Private Function CalcolaTotaleIncPagSosp(RecSetContropartite As ADODB.Recordset) As Variant
    Dim Totale      As Variant
    
    Totale = 0
    
    If Not RecSetContropartite Is Nothing Then
        If RecSetContropartite.State = adStateOpen Then
            If RecSetContropartite.RecordCount > 0 Then
                RecSetContropartite.MoveFirst
                
                While Not RecSetContropartite.EOF
                    Totale = Totale + NVL(RecSetContropartite.Fields("IMPONIBILEINCASSATO").Value, 0) + _
                                      NVL(RecSetContropartite.Fields("IMPOSTAINCASSATA").Value, 0)
                    
                    RecSetContropartite.MoveNext
                Wend
            End If
        End If
    End If
    
    CalcolaTotaleIncPagSosp = Totale
End Function

Private Sub FindAliquota(CodiceAliquota As Variant, TipoEsigIVA As Variant, PercVent As Variant)
    Dim Sql     As Variant
    Dim RecSet  As ADODB.Recordset
    
    Sql = "SELECT CG28_FLGSOSPIMP, CG28_ALIQIVAVENT FROM CG28_TABCODIVA WHERE CG28_CODICE = '" & CodiceAliquota & "'"
    
    Set RecSet = Connessione.Execute(Sql, , adCmdText)
    If Not RecSet.EOF Then
        TipoEsigIVA = NVL(RecSet.Fields("CG28_FLGSOSPIMP").Value, 0)
        PercVent = NVL(RecSet.Fields("CG28_ALIQIVAVENT").Value, 0)
    End If
    
    Set RecSet = Nothing
End Sub

Private Sub GetDatiFatturaOrigine(NumRegFattura As Variant, ByRef CG41_DATAREGIVA As Variant, ByRef CG41_DATAREG As Variant, ByRef CG41_SEZIONALE As Variant)
    Dim Sql     As Variant
    Dim RecSet  As ADODB.Recordset
    
    Sql = "SELECT CG41_DATAREGIVA, CG41_DATAREG, CG41_SEZIONALE" & _
         " FROM CG41_PRIMANOTA" & _
         " WHERE CG41_DITTA_CG18 = " & CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta & _
         " AND   CG41_NUMREG = '" & NumRegFattura & "'"

    Set RecSet = Connessione.Execute(Sql, , adCmdText)
    If Not RecSet.EOF Then
        CG41_DATAREGIVA = RecSet.Fields("CG41_DATAREGIVA").Value
        CG41_DATAREG = RecSet.Fields("CG41_DATAREG").Value
        CG41_SEZIONALE = RecSet.Fields("CG41_SEZIONALE").Value
    End If
    
    Set RecSet = Nothing
End Sub

Private Sub ScriviEcPort(RecSetScad As ADODB.Recordset, NumRegPagamento As Variant)
    Dim ClsMemPortaf            As EPBO_MEMPORTAF.CLSEP_GESTECPORT
    
    Set ClsMemPortaf = New EPBO_MEMPORTAF.CLSEP_GESTECPORT
    
    If RecSetScad Is Nothing Then
        Exit Sub
    End If
    
    If RecSetScad.State = adStateClosed Then
        Exit Sub
    End If
    
    If RecSetScad.RecordCount = 0 Then
        Exit Sub
    End If
    
    ClsMemPortaf.DittaScadenza = CallingForm.ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    Set ClsMemPortaf.CPInput.Connessione = Connessione
    
    ClsMemPortaf.ValutaPagamento = "EURO"
    ClsMemPortaf.CambioPagamento = 0
    ClsMemPortaf.NumRegPagamento = NumRegPagamento
    ClsMemPortaf.CausalePagamento = "102"
    ClsMemPortaf.DescCauPagamento = "descrizione 102"
    ClsMemPortaf.CausaleAbbuono = "105"
    ClsMemPortaf.DescCauAbbuono = "descrizione 105"
    ClsMemPortaf.DataRegPagamento = "26/01/2011"
    
    RecSetScad.MoveFirst
    While Not RecSetScad.EOF
        ClsMemPortaf.NumRegScadenza = RecSetScad.Fields("EF01_NUMREG_CO99").Value
        ClsMemPortaf.RigaContScadenza = RecSetScad.Fields("EF01_RIGACONT_CG42").Value
        ClsMemPortaf.ProgECScadenza = RecSetScad.Fields("EF01_PROGEC_EC01").Value
        ClsMemPortaf.ProgEffScadenza = RecSetScad.Fields("EF01_PROGEFF").Value
        
        ClsMemPortaf.ImportoPagato = RecSetScad.Fields("EF01_IMPEFFNETPREV_S").Value
        ClsMemPortaf.ImportoAbbuono = 0
        ClsMemPortaf.RigaContPagamento = 1 ' VALORIZZARE CON IL NUM.RIGA CONT. RELATIVO AL CLIENTE
        
        ClsMemPortaf.ChiudiScadenzaPerChiave
        If ClsMemPortaf.Stato <> 0 Then
            MsgBox "Errore in ChiudiScadenzaPerChiave: " & ClsMemPortaf.Errore
            Exit Sub
        End If
        
        RecSetScad.MoveNext
    Wend
    
    Set ClsMemPortaf = Nothing
End Sub

'
' Ritorna TRUE se la variabile ha un valore definito (cioè non è nè EMPTY nè NULL)
'
Private Function IsDefinito(Variabile As Variant) As Boolean
    On Error Resume Next
    If IsNull(Variabile) Or IsEmpty(Variabile) Or Variabile = "" Then
        IsDefinito = False
    Else
        IsDefinito = True
    End If
    Err.Clear
End Function
