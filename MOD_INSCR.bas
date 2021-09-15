Attribute VB_Name = "MOD_INSCR"
Option Explicit

Public Sub ValidarInscrEstadual(ByVal ValorInscricao As String, ByVal ValorUF As String, ByVal DestinoResp As MSForms.Label)

    Dim IE As CLS_INSCR
    Set IE = New CLS_INSCR
    
    IE.Inscricao = ValorInscricao
    
    Select Case ValorUF
        Case "PR"
            DestinoResp.Caption = IE.ValidarPR
        Case "SC"
            DestinoResp.Caption = IE.ValidarSC
        Case Else
            DestinoResp.Caption = ""
            Call MsgBox("Ainda não implementado para outras UF, somente PR e SC", vbExclamation, "MOD_INSCR_ValidarInscrEstadual")
    End Select
    
End Sub

Public Sub CarregarListaUF(ByVal CBox As MSForms.ComboBox)

    On Error GoTo TratarErro

    CBox.Clear
    With CBox
        .AddItem "RO"
        .AddItem "AC"
        .AddItem "AM"
        .AddItem "RR"
        .AddItem "PA"
        .AddItem "AP"
        .AddItem "TO"
        .AddItem "MA"
        .AddItem "PI"
        .AddItem "CE"
        .AddItem "RN"
        .AddItem "PB"
        .AddItem "AL"
        .AddItem "SE"
        .AddItem "BA"
        .AddItem "MG"
        .AddItem "ES"
        .AddItem "RJ"
        .AddItem "SP"
        .AddItem "PR"
        .AddItem "SC"
        .AddItem "RS"
        .AddItem "MS"
        .AddItem "MT"
        .AddItem "GO"
        .AddItem "DF"
    End With

    CBox.ListIndex = 19
    
    On Error GoTo 0
    
FinalizarSub:
    On Error Resume Next
    Set CBox = Nothing
    Exit Sub

TratarErro:
    Call MsgBox("Nº Erro: " & Err.Number & vbCr & "Descrição: " & Err.Description, vbCritical, "MOD_INSCR_CarregarListaUF")
    GoTo FinalizarSub
    
End Sub

Public Sub Main()
    'Deve ser chamado dentro de EstaPasta_de_trabalho método Workbook_Open
    Application.Visible = False
    Load FR_INSCRESTAD
    FR_INSCRESTAD.Show
End Sub
