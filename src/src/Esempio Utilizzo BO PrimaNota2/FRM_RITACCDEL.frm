VERSION 5.00
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.6#0"; "TMS_EDIT.ocx"
Begin VB.Form FRM_RITACCDEL 
   Caption         =   "Gestione ritenute acconto - cancellazione"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   Icon            =   "FRM_RITACCDEL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2475
   ScaleMode       =   0  'User
   ScaleWidth      =   5385.834
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BUT_REGISTRAMODIFICHE 
      Caption         =   "Registra Modifiche"
      Height          =   345
      Left            =   3270
      Picture         =   "FRM_RITACCDEL.frx":27A2
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dati relativi alla ritenuta:"
      Height          =   1065
      Left            =   60
      TabIndex        =   6
      Top             =   90
      Width           =   5115
      Begin VB.CommandButton BUT_DELETE 
         Caption         =   "Cancella mov. ritenuta"
         Height          =   345
         Left            =   3210
         Picture         =   "FRM_RITACCDEL.frx":28EC
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   600
         Width           =   1785
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DITTA 
         Height          =   300
         Left            =   1440
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
      Begin VB.Label Label1 
         Caption         =   "Ditta"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   330
         Width           =   735
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREG 
         Height          =   300
         Left            =   1440
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
      Begin VB.Label Label15 
         Caption         =   "Num. reg. ritenuta"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dati relativi al pagamento:"
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   1230
      Width           =   5115
      Begin VB.CommandButton BUT_DELETEPAG 
         Caption         =   "Cancella pag. ritenuta"
         Height          =   345
         Left            =   3210
         Picture         =   "FRM_RITACCDEL.frx":2A36
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   270
         Width           =   1785
      End
      Begin PRJFW_EDIT.TxtEdit TXT_NUMREGPAG 
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   300
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
      Begin VB.Label Label17 
         Caption         =   "Num. reg. pagam."
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   330
         Width           =   1305
      End
   End
End
Attribute VB_Name = "FRM_RITACCDEL"
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

Private Sub BUT_DELETE_Click()
    On Error GoTo Err_BUT_DELETE_Click
    
    ClsRitenute.CPInput.Ditta = TXT_DITTA.Text
    ClsRitenute.CPInput.NumeroRegistrazione = TXT_NUMREG.Text
    ClsRitenute.CPInput.ForzaCancellazioneRitenutaDaPrimaNota = True
    ClsRitenute.CPInput.ForzaCancellazionePagRitDaPrimaNota = True
    
    ClsRitenute.CancellaRitenuta
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in CancellaRitenuta"
        Exit Sub
    End If
    
Exit Sub
Err_BUT_DELETE_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_DELETE_Click"
    Exit Sub
End Sub

Private Sub BUT_DELETEPAG_Click()
    On Error GoTo Err_BUT_DELETEPAG_Click
    
    ClsRitenute.CPInput.Ditta = TXT_DITTA.Text
    ClsRitenute.CPInput.NumRegPagamento = TXT_NUMREGPAG.Text
    ClsRitenute.CPInput.ForzaCancellazionePagRitDaPrimaNota = True
    
    ClsRitenute.CancellaPagamentoRitenuta
    If ClsRitenute.Stato <> tsGestRitOk Then
        MsgBox ClsRitenute.Errore, , "Errore in CancellaPagamentoRitenuta"
        Exit Sub
    End If
    
Exit Sub
Err_BUT_DELETEPAG_Click:
    MsgBox Err.Number & " - " & Err.Description, , "BUT_DELETEPAG_Click"
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

Private Sub Form_Load()
    On Error GoTo Err_Form_Load
    
    Set ClsRitenute = New CGBO_RITENUTE.CLSCG_GESTRITENUTE
    ClsRitenute.CPInput.Sconnect = StrConnect
    
Exit Sub
Err_Form_Load:
    MsgBox Err.Number & " - " & Err.Description, , "Form_Load"
    Exit Sub
End Sub
