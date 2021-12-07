Attribute VB_Name = "MOD_INSCR"
Option Explicit

Public Sub ValidarInscrEstadual(ByVal ValorInscricao As String, ByVal ValorUF As String, ByVal DestinoResp As MSForms.Label)

    Dim IE As CLS_INSCR
    Set IE = New CLS_INSCR
    
    Let IE.Inscricao = ValorInscricao
    Let IE.Estado = ValorUF
    'Verifica o Estado selecionado pelo usuário
    'e se a quantidade de caracteres digitada é válida
    
    If (IE.ValidarDigitos > 0) Then
        DestinoResp.Caption = "Deve conter " & IE.ValidarDigitos & " caracteres"
        Exit Sub
    End If
    
    Select Case ValorUF
        Case "PR"
            DestinoResp.Caption = IE.ValidarPR
        Case "GO"
            DestinoResp.Caption = IE.ValidarGO
        Case "AP"
            DestinoResp.Caption = IE.ValidarAP
        Case "RS"
            DestinoResp.Caption = IE.ValidarRS
        Case "SP"
            DestinoResp.Caption = IE.ValidarSP
        Case "RJ"
            DestinoResp.Caption = IE.ValidarRJ
        Case "DF", "AC"
            DestinoResp.Caption = IE.ValidarDF
        Case "MG"
            DestinoResp.Caption = IE.ValidarMG
        Case "AM"
            DestinoResp.Caption = "Ainda não implementado"
        Case "BA"
            DestinoResp.Caption = "Ainda não implementado"
        Case "MT"
            DestinoResp.Caption = "Ainda não implementado"
        Case "PE"
            DestinoResp.Caption = "Ainda não implementado"
        Case "RN"
            DestinoResp.Caption = "Ainda não implementado"
        Case "RO"
            DestinoResp.Caption = "Ainda não implementado"
        Case "RR"
            DestinoResp.Caption = "Ainda não implementado"
        Case "SE"
            DestinoResp.Caption = "Ainda não implementado"
        Case "TO"
            DestinoResp.Caption = "Ainda não implementado"
        Case Else
            DestinoResp.Caption = IE.ValidarSC
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

Public Sub MostrarExcel()
    Application.Visible = True
End Sub
