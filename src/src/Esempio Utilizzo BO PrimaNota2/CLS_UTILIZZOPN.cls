VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_UTILIZZOPN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Interfaccia standard
Implements Cinterface

'Variabili usate per implentare le proprietà standard dell'interfaccia
Private Gbol_IsActive                   As Boolean
Private GbolIsCalled                    As Boolean
Private Glng_Top                        As Long
Private Glng_Left                       As Long
Private Gbol_CanInsert                  As Boolean
Private Gbol_CanDelete                  As Boolean
Private Gbol_CanUpdate                  As Boolean
Private Gbol_CanQbe                     As Boolean
Private Gint_WinState                   As Integer
Private Gbol_WinModal                   As Boolean
Private Gint_Mode                       As Integer
Private Gstr_MenuModulo                 As Variant
Private Gstr_MenuPadre                  As Variant
Private Gstr_MenuProg                   As Variant
Private Gstr_Caption                    As Variant
Private Glng_Parent                     As Long
Private Gobj_StatusBar                  As Object
Private Gobj_ActiveFrame                As Object
Private Gobj_ActiveDll                  As Cinterface
Private GclsVoceMenu                    As CLSFW_VOCE
Private Gcls_Global                     As CLSFW_Global
Private WithEvents Gobj_Window          As CLSFW_Window
Attribute Gobj_Window.VB_VarHelpID = -1
Private WithEvents ActiveNavigator      As CLSFW_Navigator
Attribute ActiveNavigator.VB_VarHelpID = -1
Private Gcon_Connect                    As ADODB.Connection

'Riferimento alla form
Private FormMain                        As FRM_MAIN

'Integrazione multimediale
Private WithEvents CHyperMedia          As HMUO_HYPERMEDIA.CLSHM_MAIN
Attribute CHyperMedia.VB_VarHelpID = -1

'Evento di ritorno dati a programmi chiamanti
Public Event DataTransfer(ByVal Data As Variant)

'Inizializzazione classe principale
Private Sub Class_Initialize()
    On Error Resume Next
    Set Gobj_ActiveDll = Me
    Err.Clear
End Sub

'Terminazione classe principale
Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    'Distruggo gli oggetti della classe
    DestroyClass
    
Exit Sub
ErrorHandler:
    Err.Clear
    MsgBox "Errore VB n." & Err.Number & " - " & Err.Description & " nel Terminate della classe principale"
End Sub


'Proprietà interfaccia standard
Private Property Set Cinterface_ActiveDll(RHS As FWBO_LibSrv.Cinterface)
    On Error Resume Next
    Set Gobj_ActiveDll = RHS
    Err.Clear
End Property
Private Property Get Cinterface_ActiveDll() As FWBO_LibSrv.Cinterface
    On Error Resume Next
    Set Cinterface_ActiveDll = Gobj_ActiveDll
    Err.Clear
End Property

Private Property Get Cinterface_ActiveFrame() As Object
    On Error Resume Next
    Set Cinterface_ActiveFrame = Gobj_ActiveFrame
    Err.Clear
End Property
Private Property Set Cinterface_ActiveFrame(RHS As Object)
    On Error Resume Next
    Set Gobj_ActiveFrame = RHS
    Err.Clear
End Property

Private Property Set Cinterface_ActiveNavigator(RHS As FWBO_LibSrv.CLSFW_Navigator)
    On Error Resume Next
    Set ActiveNavigator = RHS
    Err.Clear
End Property
Private Property Get Cinterface_ActiveNavigator() As FWBO_LibSrv.CLSFW_Navigator
    On Error Resume Next
    Set Cinterface_ActiveNavigator = ActiveNavigator
    Err.Clear
End Property

Private Property Let Cinterface_CanDelete(ByVal RHS As Boolean)
    On Error Resume Next
    Gbol_CanDelete = RHS
    Err.Clear
End Property
Private Property Get Cinterface_CanDelete() As Boolean
    On Error Resume Next
    Cinterface_CanDelete = Gbol_CanDelete
    Err.Clear
End Property

Private Property Let Cinterface_CanInsert(ByVal RHS As Boolean)
    On Error Resume Next
    Gbol_CanInsert = RHS
    Err.Clear
End Property
Private Property Get Cinterface_CanInsert() As Boolean
    On Error Resume Next
    Cinterface_CanInsert = Gbol_CanInsert
    Err.Clear
End Property

Private Property Let Cinterface_CanQbe(ByVal RHS As Boolean)
    On Error Resume Next
    Gbol_CanQbe = RHS
    Err.Clear
End Property
Private Property Get Cinterface_CanQbe() As Boolean
    On Error Resume Next
    Cinterface_CanQbe = Gbol_CanQbe
    Err.Clear
End Property

Private Property Let Cinterface_CanUpdate(ByVal RHS As Boolean)
    On Error Resume Next
    Gbol_CanUpdate = RHS
    Err.Clear
End Property
Private Property Get Cinterface_CanUpdate() As Boolean
    On Error Resume Next
    Cinterface_CanUpdate = Gbol_CanUpdate
    Err.Clear
End Property

Private Property Let Cinterface_Caption(ByVal RHS As Variant)
    On Error Resume Next
    Gstr_Caption = RHS
    Err.Clear
End Property
Private Property Get Cinterface_Caption() As Variant
    On Error Resume Next
    Cinterface_Caption = Gstr_Caption
    Err.Clear
End Property

Private Property Get Cinterface_ClsGlobal() As FWBO_LibSrv.CLSFW_Global
    On Error Resume Next
    Set Cinterface_ClsGlobal = Gcls_Global
    Err.Clear
End Property
Private Property Set Cinterface_ClsGlobal(RHS As FWBO_LibSrv.CLSFW_Global)
    On Error Resume Next
    Set Gcls_Global = RHS
    Err.Clear
End Property

Private Property Get Cinterface_Connection() As ADODB.Connection
    On Error Resume Next
    Set Cinterface_Connection = Gcon_Connect
    Err.Clear
End Property
Private Property Set Cinterface_Connection(RHS As ADODB.Connection)
    On Error Resume Next
    Set Gcon_Connect = RHS
    Err.Clear
End Property

Private Property Get Cinterface_Cwindow() As FWBO_LibSrv.CLSFW_Window
    On Error Resume Next
    Set Cinterface_Cwindow = Gobj_Window
    Err.Clear
End Property

Private Property Let Cinterface_Filtro(ByVal RHS As Variant)
End Property
Private Property Get Cinterface_Filtro() As Variant
End Property

Private Property Let Cinterface_hwndParent(ByVal RHS As Long)
    On Error Resume Next
    Glng_Parent = RHS
    Err.Clear
End Property
Private Property Get Cinterface_hwndParent() As Long
    On Error Resume Next
    Cinterface_hwndParent = Glng_Parent
    Err.Clear
End Property

Private Property Let Cinterface_IsCalled(ByVal RHS As Boolean)
   On Error Resume Next
   GbolIsCalled = RHS
   Err.Clear
End Property
Private Property Get Cinterface_IsCalled() As Boolean
   On Error Resume Next
   Cinterface_IsCalled = GbolIsCalled
    Err.Clear
End Property

Private Property Let Cinterface_MenuModulo(ByVal RHS As Variant)
    On Error Resume Next
    Gstr_MenuModulo = RHS
    Err.Clear
End Property
Private Property Get Cinterface_MenuModulo() As Variant
    On Error Resume Next
    Cinterface_MenuModulo = Gstr_MenuModulo
    Err.Clear
End Property

Private Property Let Cinterface_MenuPadre(ByVal RHS As Variant)
    On Error Resume Next
    Gstr_MenuPadre = RHS
    Err.Clear
End Property
Private Property Get Cinterface_MenuPadre() As Variant
    On Error Resume Next
    Cinterface_MenuPadre = Gstr_MenuPadre
    Err.Clear
End Property

Private Property Let Cinterface_MenuProg(ByVal RHS As Variant)
    On Error Resume Next
    Gstr_MenuProg = RHS
    Err.Clear
End Property
Private Property Get Cinterface_MenuProg() As Variant
    On Error Resume Next
    Cinterface_MenuProg = Gstr_MenuProg
    Err.Clear
End Property

Private Property Let Cinterface_ProgramMode(ByVal RHS As FWBO_LibSrv.Mode)
    On Error Resume Next
    Gint_Mode = RHS
    Err.Clear
End Property
Private Property Get Cinterface_ProgramMode() As FWBO_LibSrv.Mode
    On Error Resume Next
    Cinterface_ProgramMode = Gint_Mode
    Err.Clear
End Property

Private Property Set Cinterface_StatusBar(RHS As Object)
    On Error Resume Next
    Set Gobj_StatusBar = RHS
    Err.Clear
End Property
Private Property Get Cinterface_StatusBar() As Object
    On Error Resume Next
    Set Cinterface_StatusBar = Gobj_StatusBar
    Err.Clear
End Property

Private Property Get Cinterface_WindowModal() As Boolean
    On Error Resume Next
    Cinterface_WindowModal = Gbol_WinModal
    Err.Clear
End Property
Private Property Let Cinterface_WindowModal(ByVal RHS As Boolean)
    On Error Resume Next
    Gbol_WinModal = RHS
    Err.Clear
End Property

Private Property Let Cinterface_WindowState(ByVal RHS As FWBO_LibSrv.WindowState)
    On Error Resume Next
    Gint_WinState = RHS
    Err.Clear
End Property
Private Property Get Cinterface_WindowState() As FWBO_LibSrv.WindowState
    On Error Resume Next
    Cinterface_WindowState = Gint_WinState
    Err.Clear
End Property

Private Property Let Cinterface_Top(ByVal RHS As Variant)
    On Error Resume Next
    Glng_Top = RHS
    Err.Clear
End Property
Private Property Get Cinterface_Top() As Variant
    On Error Resume Next
    Cinterface_Top = Glng_Top
    Err.Clear
End Property

Private Property Let Cinterface_Left(ByVal RHS As Variant)
    On Error Resume Next
    Glng_Left = RHS
    Err.Clear
End Property
Private Property Get Cinterface_Left() As Variant
    On Error Resume Next
    Cinterface_Left = Glng_Left
    Err.Clear
End Property

Private Property Let Cinterface_IsActive(ByVal RHS As Boolean)
    On Error Resume Next
    Gbol_IsActive = RHS
    Err.Clear
End Property
Private Property Get Cinterface_IsActive() As Boolean
    On Error Resume Next
    Cinterface_IsActive = Gbol_IsActive
    Err.Clear
End Property

Private Property Set Cinterface_ClsVoceMenu(RHS As FWBO_GESMODULI.CLSFW_VOCE)
    On Error Resume Next
    Set GclsVoceMenu = RHS
    Err.Clear
End Property
Private Property Get Cinterface_ClsVoceMenu() As FWBO_GESMODULI.CLSFW_VOCE
    On Error Resume Next
    Set Cinterface_ClsVoceMenu = GclsVoceMenu
    Err.Clear
End Property


'Metodi interfaccia standard
Private Function Cinterface_CloseForm() As Variant
    On Error Resume Next
    
    'Forzo la chiusura della form alla chiusura dello startup
    Cinterface_CloseForm = False
    Unload FormMain
    If FormMain.ActiveInterface Is Nothing Then
        Cinterface_CloseForm = True
    End If
    
    Err.Clear
End Function

Private Function Cinterface_ShowForm() As Variant
    On Error Resume Next
    
    'Istanzio l'oggetto window (ossia il container della form)
    Set Gobj_Window = New CLSFW_Window
    
    'Istanzio la form
    Set FormMain = New FRM_MAIN
    Set FormMain.ActiveInterface = Gobj_ActiveDll
    Set FormMain.ActiveClass = Me
    'Subclasso la form
    Set Gobj_ActiveDll.Cwindow.HookedForm = FormMain
    
    'Imposto la modalità di apertura programma (Modal/Modless)
    If Gbol_WinModal = False Then
       Gobj_Window.hWnd = FormMain.hWnd
       Gobj_Window.Parent = Glng_Parent
       FormMain.Caption = Gstr_Caption
       FormMain.Show 0
    Else
       FormMain.Caption = Gstr_Caption
       FormMain.Show 1
    End If
    
    Err.Clear
End Function

Private Function Cinterface_TransferData(ByVal Data As Variant) As Variant
    On Error Resume Next
    RaiseEvent DataTransfer(Data)
    Err.Clear
End Function

'Refresh del navigator
Private Sub Gobj_Window_ActivateApp()
    On Error Resume Next
    Set Gobj_ActiveDll.ClsGlobal.ActiveInterface = Gobj_ActiveDll
    Gobj_ActiveDll.ActiveNavigator.SetStatus (Gobj_ActiveDll.ActiveFrame.Status)
    Gobj_ActiveDll.ActiveNavigator.Refresh
    Err.Clear
End Sub

'Export to office
Private Sub ActiveNavigator_BeforeExportToOffice(Cancel As Boolean, Pcls_ExpToOffice As Object)
    Dim Pvar_Decode      As Variant
    Dim Pvar_Valori      As Variant
   
    On Error Resume Next
    
    If Not Gobj_ActiveDll.IsActive Then
        Exit Sub
    End If
    
    Pcls_ExpToOffice.TitoloMaster = "Esempio prima nota"
    
    Err.Clear
End Sub

'Personalizzazione layout
Private Sub ActiveNavigator_InitializeLayoutObjectModel(Cancel As Boolean, ClsLayout As PRJFW_PRSLAYOUT.CLSFW_PRSLAYOUT)
    On Error Resume Next
    
    If Not Gobj_ActiveDll.IsActive Then
       Exit Sub
    End If
    ClsLayout.AddStruttura "Esempio prima nota", FormMain, False, False, PRJFW_PRSLAYOUT.tsFormGestione
    
    Err.Clear
End Sub

'Scripting
Private Sub ActiveNavigator_InitializeScriptObjectModel(Cancel As Boolean, ClsScript As FWUO_TMSDEVELOP.CLSFW_PRSVBSCRIPT)
    On Error Resume Next
    
    If Not Gobj_ActiveDll.IsActive Then
       Exit Sub
    End If
    ClsScript.AddObject FormMain.Name, FormMain, 0, "Esempio prima nota"
    Gobj_ActiveDll.ActiveFrame.Name = ""
    ClsScript.AddObject Gobj_ActiveDll.ActiveFrame.Name, Gobj_ActiveDll.ActiveFrame, FWUO_TMSDEVELOP.tsDataSource, Gobj_ActiveDll.ActiveFrame.Name
    
    Err.Clear
End Sub

'Integrazione con HyperMedia
Private Sub CHyperMedia_AfterMediaCreate(Cancel As Boolean, CodiceMedia As Variant)
    On Error Resume Next
    
    Cancel = True
    
    Err.Clear
End Sub

Private Sub ActiveNavigator_ButtonPressed(Cancel As Boolean, ButtonPress As FWBO_LibSrv.ButtonIndex)
    Dim PvarIDMedia As Variant

    On Error Resume Next

    If Not Gobj_ActiveDll.IsActive Then
       Exit Sub
    End If

    If ButtonPress = tsHyperMedia Then
       If Gobj_ActiveDll.ActiveFrame Is Nothing Then
          Exit Sub
       End If
       If Gobj_ActiveDll.ActiveFrame.Recordset Is Nothing Then
          Exit Sub
       End If
       If Gobj_ActiveDll.ActiveFrame.Status = tsInsert Then
          MsgBox "Il collegamento ai dati multimediali è possibile solamente in fase di variazione", vbInformation, "Avviso Utente"
          Exit Sub
       ElseIf Gobj_ActiveDll.ActiveFrame.Status = tsModify Then
          Gobj_ActiveDll.ActiveFrame.Update
       End If
    End If
    
    Err.Clear
End Sub

Public Property Get ObjApplication() As Object
    On Error Resume Next
    Set ObjApplication = App
    Err.Clear
End Property

Private Sub DestroyClass()
    On Error GoTo ErrorHandler
    
    'Distruggo gli oggetti della classe
    Set Gobj_StatusBar = Nothing
    Set Gobj_ActiveFrame = Nothing
    Set Gobj_ActiveDll = Nothing
    Set GclsVoceMenu = Nothing
    Set Gcls_Global = Nothing
    Set Gobj_Window = Nothing
    Set ActiveNavigator = Nothing
    Set CHyperMedia = Nothing
    Set Gcon_Connect = Nothing
    Set FormMain = Nothing
    
Exit Sub
ErrorHandler:
    Err.Clear
    MsgBox "Errore VB n." & Err.Number & " - " & Err.Description & " nel metodo DestroyClass della classe principale"
End Sub
