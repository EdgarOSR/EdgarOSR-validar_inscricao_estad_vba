VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FR_INSCRESTAD 
   Caption         =   "VERIFICADOR INSCRIÇÕES ESTADUAIS BRASIL"
   ClientHeight    =   1980
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4740
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
    If (Me.C_INSCR.Value <> "") Then
        Call MOD_INSCR.ValidarInscrEstadual(Me.C_INSCR.Value, Me.C_UF.Value, Me.L_RESULTADO)
    Else
        Me.L_RESULTADO = "Digite um valor"
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.C_INSCR.Value = ""
    Me.C_INSCR.SetFocus
    Call MOD_INSCR.CarregarListaUF(C_UF)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Let Cancel = vbTrue
    End If
End Sub
