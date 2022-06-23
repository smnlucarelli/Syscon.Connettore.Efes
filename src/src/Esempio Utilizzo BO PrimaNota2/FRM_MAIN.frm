VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_MAIN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Utilizzo business object prima nota"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "FRM_MAIN.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   9900
      Width           =   1245
   End
   Begin VB.CommandButton CMD_ESCI 
      Caption         =   "Esci"
      Height          =   345
      Left            =   3960
      TabIndex        =   1
      Top             =   9930
      Width           =   945
   End
   Begin MSComctlLib.ImageList IML_IMAGELIST 
      Left            =   3960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRM_MAIN.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRM_MAIN.frx":28FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TVW_MENU 
      Height          =   9795
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   17277
      _Version        =   393217
      Indentation     =   617
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "IML_IMAGELIST"
      Appearance      =   1
   End
End
Attribute VB_Name = "FRM_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private StrConnect          As Variant
Private Connessione         As ADODB.Connection

'Riferimento all'interfaccia standard della classe
Public ActiveInterface      As Cinterface

'Riferimento all'interfaccia estesa della classe
Public ActiveClass          As CLS_UTILIZZOPN

Private Sub CMD_ESCI_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
'    Dim ClsPoolPrimaNota    As CLSCG_POOLPNOTA
'    Dim Filtro              As Variant
'    Dim Causale             As Variant
'
'    Set ClsPoolPrimaNota = New CLSCG_POOLPNOTA
'    ClsPoolPrimaNota.PbolActiveQuery = True
'    Set ClsPoolPrimaNota.ActiveInterfaceQuery = ActiveInterface
'    Set ClsPoolPrimaNota.Connessione = Connessione
'
'    Filtro = "CG41_NUMREG = '201200014576'"
'    Causale = "1"
'
'    ClsPoolPrimaNota.Causale = Causale
'    ClsPoolPrimaNota.Filtro = Filtro
'    ClsPoolPrimaNota.ExecutePNota
    
    Dim Sql     As Variant
    Dim RecSet  As ADODB.Recordset
    
    Sql = " SELECT *," & vbCrLf & _
          "        CAST(0 AS DECIMAL(13,2)) AS PAGATO" & vbCrLf & _
          " INTO #EF01_TEMP" & vbCrLf & _
          " FROM EF01_SCADENZE WITH (NOLOCK)" & vbCrLf & _
          " WHERE 1 = 2"
    
    Connessione.Execute Sql, , adCmdText
    
    Sql = " SELECT *" & vbCrLf & _
          " FROM #EF01_TEMP"
    
    Set RecSet = New ADODB.Recordset
    Set RecSet.ActiveConnection = Connessione
    RecSet.CursorLocation = adUseClient
    RecSet.CursorType = adOpenDynamic
    RecSet.LockType = adLockBatchOptimistic
    RecSet.Open Sql
    
    Set RecSet.ActiveConnection = Nothing
    
    RecSet.AddNew
    RecSet.Fields("EF01_DITTA_CG18").Value = 80
    RecSet.Fields("PAGATO").Value = 123.45
    RecSet.Update
    
End Sub

'Private Sub Command1_Click()
'    Dim ClsCGConnect As Object
'
'    Set ClsCGConnect = CreateObject("CGBO_LOOKUPDECODE.CLSCG_CONNECT")
'
'    Set ClsCGConnect.ActiveInterface = ActiveInterface
'    ClsCGConnect.CallAnagraficaBeniUsati
'
'    ClsCGConnect.TerminateConnect
'    Set ClsCGConnect.ActiveInterface = Nothing
'    Set ClsCGConnect = Nothing
'End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    TVW_MENU.SetFocus
    
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Me.Left = 100
    Me.Top = 100
    
    StrConnect = ActiveInterface.ClsGlobal.Gcls_LibConnect.GetExtendedProperties
    
    '
    ' Creo la connessione
    '
    Set Connessione = New ADODB.Connection
    Connessione.ConnectionString = StrConnect
    Connessione.CursorLocation = adUseClient
    Connessione.Open
    
    CaricaMenu
    
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim UserObject As Variant
    
    On Error GoTo Err_Form_Unload
    
    ActiveInterface.ClsGlobal.RemoveCurrentInterface ActiveInterface
    
    '
    ' Rimuovo i riferimenti alla classe per la personalizzazione del layout e dello script
    '
    UserObject = ActiveInterface.ClsVoceMenu.Classe
    If Not ActiveInterface.ActiveNavigator.ClsScript Is Nothing Then
        ActiveInterface.ActiveNavigator.ClsScript.TerminateByUserObject UserObject
    End If
    Set ActiveInterface.ActiveNavigator.ClsLayout = Nothing
    Set ActiveInterface.ActiveNavigator.ClsScript = Nothing
    
    '
    ' Distruggo l'ActiveInterface
    '
    Set ActiveInterface.ClsGlobal.ActiveInterface = Nothing
    Set ActiveInterface.ClsGlobal.CallInterface = Nothing
    Set ActiveInterface.ActiveDll = Nothing
    Set ActiveInterface.ActiveNavigator.ActiveInterface = Nothing
    Set ActiveInterface.ActiveDll = Nothing
    Set ActiveInterface = Nothing
    
    '
    ' Distruggo la classe principale
    '
    Set ActiveClass = Nothing
Exit Sub
Err_Form_Unload:
    MsgBox Err.Number & " - " & Err.Description & " in Form_Unload"
    Err.Clear
End Sub

Private Sub CaricaMenu()
    Dim Nodo        As Node
    On Error GoTo Err_CaricaMenu
    
    TVW_MENU.Nodes.Clear
    
    Set Nodo = TVW_MENU.Nodes.Add(, , "K_NEW_BO", "Test nuovo b.o. GestRegPN", 1, 1)
    Set Nodo = TVW_MENU.Nodes.Add("K_NEW_BO", tvwChild, "K_MOVCONT_INS_REPEAT", "Inserimento / variazione", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_NEW_BO", tvwChild, "K_MOVIVA_INS_REPEAT", "Inserimento mov. IVA massivo", 2, 2)
    
    Set Nodo = TVW_MENU.Nodes.Add(, , "K_MOVCONT_CASSA", "Movimenti contabili - diversi a diversi - CASSA", 1, 1)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVCONT_CASSA", tvwChild, "K_MOVCONT_INS_CASSA", "Inserimento", 2, 2)
    
    Set Nodo = TVW_MENU.Nodes.Add(, , "K_MOVCONT", "Movimenti contabili - diversi a diversi", 1, 1)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVCONT", tvwChild, "K_MOVCONT_INS", "Inserimento", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVCONT", tvwChild, "K_MOVCONT_UPD", "Modifica", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVCONT", tvwChild, "K_MOVCONT_DEL", "Cancellazione", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVCONT", tvwChild, "K_MOVCONT_INS_ECPORT", "Inserimento con EC/Port", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVCONT", tvwChild, "K_MOVCONT_INS_ABBUONO", "Inserimento con abbuono", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVCONT", tvwChild, "K_MOVCONT_UPD_ABBUONO", "Modifica con abbuono", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVCONT", tvwChild, "K_MOVCONT_INS_ECPORT_MULTIPLO", "Inserimento con EC/Port multiplo", 2, 2)
    
    Set Nodo = TVW_MENU.Nodes.Add(, , "K_MOVIVA", "Movimenti IVA", 1, 1)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVIVA", tvwChild, "K_MOVIVA_INS", "Inserimento", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVIVA", tvwChild, "K_MOVIVA_UPD", "Modifica", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVIVA", tvwChild, "K_MOVIVA_DEL", "Cancellazione", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVIVA", tvwChild, "K_MOVIVA_INS_NC", "Inserimento nota credito", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVIVA", tvwChild, "K_MOVIVA_INS_AUTOFT", "Inserimento reverse charge", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVIVA", tvwChild, "K_MOVIVA_INS_IVAAUTO", "Inserimento - riga IVA automatica", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVIVA", tvwChild, "K_MOVIVA_INS_IVAAUTO_REPEAT", "Inserimento - riga IVA automatica - ins. massivo", 2, 2)
    
    Set Nodo = TVW_MENU.Nodes.Add(, , "K_INCSOSP", "Incasso fatture sospese", 1, 1)
    Set Nodo = TVW_MENU.Nodes.Add("K_INCSOSP", tvwChild, "K_INCSOSP_INS", "Inserimento", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_INCSOSP", tvwChild, "K_INCSOSP_BO", "Inserimento incasso scad. iva diff.", 2, 2)
    
    Set Nodo = TVW_MENU.Nodes.Add(, , "K_CORRG", "Corrispettivi giornalieri", 1, 1)
    Set Nodo = TVW_MENU.Nodes.Add("K_CORRG", tvwChild, "K_CORRG_INS", "Inserimento", 2, 2)
    
    Set Nodo = TVW_MENU.Nodes.Add(, , "K_RITACC", "Ritenute", 1, 1)
    Set Nodo = TVW_MENU.Nodes.Add("K_RITACC", tvwChild, "K_RITACC_INS", "Inserimento", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_RITACC", tvwChild, "K_RITACC_UPD", "Modifica", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_RITACC", tvwChild, "K_RITACC_DEL", "Cancellazione", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_RITACC", tvwChild, "K_RITACC_PAG", "Pagamento", 2, 2)
    
    Set Nodo = TVW_MENU.Nodes.Add(, , "K_MOVBENIUS", "Movimenti IVA - Beni usati", 1, 1)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVBENIUS", tvwChild, "K_MOVBENIUS_INS", "Inserimento", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVBENIUS", tvwChild, "K_MOVBENIUS_UPD", "Modifica", 2, 2)
    Set Nodo = TVW_MENU.Nodes.Add("K_MOVBENIUS", tvwChild, "K_MOVBENIUS_DEL", "Cancellazione", 2, 2)
    
    Set Nodo = TVW_MENU.Nodes.Add(, , "K_ECPORT", "Chiusura scedenze", 1, 1)
    Set Nodo = TVW_MENU.Nodes.Add("K_ECPORT", tvwChild, "K_ECPORT_CHIU_NUMPART", "Chiusura per numero partita", 2, 2)
    
    For Each Nodo In TVW_MENU.Nodes
        Nodo.Expanded = True
    Next
    
Exit Sub
Err_CaricaMenu:
    MsgBox Err.Number & " - " & Err.Description & " in CaricaMenu"
    Err.Clear
End Sub

Private Sub TVW_MENU_DblClick()
    Dim Nodo            As Node
    Dim ChiaveNodo      As Variant
    Dim ObjForm         As Object
    
    On Error GoTo Err_TVW_MENU_DblClick
    
    Set Nodo = TVW_MENU.SelectedItem
    ChiaveNodo = Nodo.Key
    
    Select Case ChiaveNodo
        Case "K_MOVCONT_INS"
            Set ObjForm = New FRM_REGMOVINS
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVCONT_INS_CASSA"
            Set ObjForm = New FRM_REGMOVINS_CASSA
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVCONT_INS_ECPORT"
            Set ObjForm = New FRM_REGMOVINS_ECPORT
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVCONT_INS_ABBUONO"
            Set ObjForm = New FRM_REGMOVINS_ABBUONO
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVCONT_UPD_ABBUONO"
            Set ObjForm = New FRM_REGMOVUPD_ABBUONO
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVCONT_UPD"
            Set ObjForm = New FRM_REGMOVUPD
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVCONT_DEL"
            Set ObjForm = New FRM_REGMOVDEL
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVIVA_INS"
            Set ObjForm = New FRM_MOVIVAINS
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVIVA_UPD"
            Set ObjForm = New FRM_MOVIVAUPD
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVIVA_DEL"
            Set ObjForm = New FRM_MOVIVADEL
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_RITACC_INS"
            Set ObjForm = New FRM_RITACCINS
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_RITACC_UPD"
            Set ObjForm = New FRM_RITACCUPD
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_RITACC_DEL"
            Set ObjForm = New FRM_RITACCDEL
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_RITACC_PAG"
            Set ObjForm = New FRM_RITACCPAG
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVBENIUS_INS"
            Set ObjForm = New FRM_MOVBENIUSINS
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVBENIUS_DEL"
            Set ObjForm = New FRM_MOVBENIUSDEL
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVBENIUS_UPD"
            Set ObjForm = New FRM_MOVBENIUSUPD
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_ECPORT_CHIU_NUMPART"
            Set ObjForm = New FRM_CHIUDISCADENZE
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
        
        Case "K_CORRG_INS"
            Set ObjForm = New FRM_CORRGINS
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
        
        Case "K_MOVIVA_INS_NC"
            Set ObjForm = New FRM_MOVIVAINS_NC
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVIVA_INS_AUTOFT"
            Set ObjForm = New FRM_MOVIVAINS_AUTOFT
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_INCSOSP_INS"
            Set ObjForm = New FRM_INCSOSP
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVIVA_INS_IVAAUTO"
            Set ObjForm = New FRM_MOVIVAINS_IVAAUTO
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVIVA_INS_IVAAUTO_REPEAT"
            Set ObjForm = New FRM_MOVIVAINS_IVAAUTO_REPEAT
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVCONT_INS_REPEAT"
            Set ObjForm = New FRM__NEW_BO_PN
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
           
        Case "K_INCSOSP_BO"
            Set ObjForm = New FRM_INCASSO_SOSP
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVCONT_INS_ECPORT_MULTIPLO"
            Set ObjForm = New FRM_REGMOVINS_ECPORT_MULTIPLO
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
        Case "K_MOVIVA_INS_REPEAT"
            Set ObjForm = New FRM_MOVIVA_NEW_BO
            ObjForm.StrConnect = StrConnect
            Set ObjForm.CallingForm = Me
            ObjForm.Show vbModal
            
    End Select
    
    Set ObjForm = Nothing
Exit Sub
Err_TVW_MENU_DblClick:
    MsgBox Err.Number & " - " & Err.Description & " in TVW_MENU_DblClick"
    Err.Clear
End Sub
