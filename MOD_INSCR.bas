Attribute VB_Name = "MOD_INSCR"
Option Explicit

Public Sub ValidarInscrEstadual(ByVal ValorInscricao As String, ByVal ValorUF As String, ByVal DestinoResp As MSForms.Label)

    Dim IE As CLS_INSCR
    Set IE = New CLS_INSCR
    
    Let IE.Inscricao = ValorInscricao
    Let IE.Estado = ValorUF
    'Verifica o Estado selecionado pelo usuário
    'e se a quantidade de caracteres digitada é válida
    
    If (IE.ValidarDigitos <> "0") Then
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
        Case "BA"
            DestinoResp.Caption = IE.ValidarBA
        Case "MT"
            DestinoResp.Caption = IE.ValidarMT
        Case "PE"
            DestinoResp.Caption = IE.ValidarPE
        Case "RN"
            DestinoResp.Caption = IE.ValidarRN
        Case "RO"
            DestinoResp.Caption = IE.ValidarRO
        Case "RR"
            DestinoResp.Caption = IE.ValidarRR
        Case "TO"
            DestinoResp.Caption = IE.ValidarTO
        Case Else
            DestinoResp.Caption = IE.ValidarSC
    End Select
End Sub

Public Sub CarregarListaUF(ByVal CBox As MSForms.ComboBox)

    On Error GoTo TratarErro

    CBox.Clear
    With CBox
        .AddItem "AC"
        .AddItem "AL"
        .AddItem "AM"
        .AddItem "AP"
        .AddItem "BA"
        .AddItem "CE"
        .AddItem "DF"
        .AddItem "ES"
        .AddItem "GO"
        .AddItem "MA"
        .AddItem "MG"
        .AddItem "MS"
        .AddItem "MT"
        .AddItem "PA"
        .AddItem "PB"
        .AddItem "PE"
        .AddItem "PI"
        .AddItem "PR"
        .AddItem "RJ"
        .AddItem "RN"
        .AddItem "RO"
        .AddItem "RR"
        .AddItem "RS"
        .AddItem "SC"
        .AddItem "SE"
        .AddItem "SP"
        .AddItem "TO"
    End With

    CBox.ListIndex = 17
    
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

Public Function TESTEX(ByVal valor As String) As String
    valor = Left$(valor, 10)
    TESTEX = Replace(valor, Mid$(valor, 3, 2), "")
End Function
