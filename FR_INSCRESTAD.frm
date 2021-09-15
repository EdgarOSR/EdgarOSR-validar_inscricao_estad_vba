VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FR_INSCRESTAD 
   Caption         =   "VERIFICADOR INSCR ESTADUAL PR/SC"
   ClientHeight    =   1980
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4728
   OleObjectBlob   =   "FR_INSCRESTAD.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FR_INSCRESTAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_FECHAR_Click()
    
    Dim Resp As Integer
    Resp = MsgBox("Deseja realmente fechar a aplicação?", vbYesNo, "Sair")
    
    If Resp = vbYes Then
        Me.Hide
        Unload Me
        Application.Visible = True
        Call ActiveWorkbook.Close(False)
    End If
    
End Sub

Private Sub B_LIMPAR_Click()
    Me.C_INSCR.Value = ""
    Me.C_INSCR.SetFocus
    Me.C_UF.Value = "PR"
    Me.L_RESULTADO = ""
End Sub

Private Sub B_VERIFICAR_Click()
    Call MOD_INSCR.ValidarInscrEstadual(Me.C_INSCR.Value, Me.C_UF.Value, Me.L_RESULTADO)
End Sub

Private Sub UserForm_Initialize()
    Me.C_INSCR.Value = ""
    Me.C_INSCR.SetFocus
    Call MOD_INSCR.CarregarListaUF(C_UF)
End Sub
