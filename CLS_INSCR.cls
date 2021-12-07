VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLS_INSCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pInscricao As String
Private pEstado As String
'Utilizar a referência Microsoft VBScript Regular Expressions 5.5

Private Sub Class_Initialize()
    pInscricao = ""
    pEstado = ""
End Sub

Property Get Inscricao() As String
    Inscricao = pInscricao
End Property

Property Let Inscricao(ByVal valorTextbox As String)
    'Expressão regular para obter somente os números do texto
    'Subtituição dos números por vazio
    Dim rgx As New RegExp
    rgx.Global = True
    rgx.Pattern = "\D"
    pInscricao = rgx.Replace(valorTextbox, "")
    Set rgx = Nothing
End Property

Property Get ValidarDigitos() As Integer
    Select Case pEstado
        Case "RO"
            If (Len(pInscricao) <> 14) Then ValidarDigitos = 14 Else ValidarDigitos = 0
        Case "DF", "MG", "AC"
            If (Len(pInscricao) <> 13) Then ValidarDigitos = 13 Else ValidarDigitos = 0
        Case "SP"
            If (Len(pInscricao) <> 12) Then ValidarDigitos = 12 Else ValidarDigitos = 0
        Case "MG", "MT", "TO"
            If (Len(pInscricao) <> 11) Then ValidarDigitos = 11 Else ValidarDigitos = 0
        Case "PR", "RS"
            If (Len(pInscricao) <> 10) Then ValidarDigitos = 10 Else ValidarDigitos = 0
        Case "RJ", "BA"
            If (Len(pInscricao) <> 8) Then ValidarDigitos = 8 Else ValidarDigitos = 0
        Case "RN"
            If ((Len(pInscricao) > 8) And (Len(pInscricao) < 11)) Then ValidarDigitos = 10 Else ValidarDigitos = 0
        Case Else
            If (Len(pInscricao) <> 9) Then ValidarDigitos = 9 Else ValidarDigitos = 0
    End Select
End Property

Property Let Estado(ByVal valorTextbox As String)
    pEstado = valorTextbox
End Property

Property Get ValidarSC() As String
    'Segundo instruções do Sintegra:
    'http://www.sintegra.gov.br/Cad_Estados/cad_SC.html
    'http://www.sintegra.gov.br/Cad_Estados/cad_ES.html
    'http://www.sintegra.gov.br/Cad_Estados/cad_CE.html
    'http://www.sintegra.gov.br/Cad_Estados/cad_MA.html
    'http://www.sintegra.gov.br/Cad_Estados/cad_MS.html
    'http://www.sintegra.gov.br/Cad_Estados/cad_PA.html
    'http://www.sintegra.gov.br/Cad_Estados/cad_PB.html
    'http://www.sintegra.gov.br/Cad_Estados/cad_AL.html
    Dim ieNum As String, origDV As String, peso As String, ieDV As String
    Dim pos As Integer, soma As Integer, resto As Integer
    
    Let ieNum = Left$(pInscricao, 8)
    Let origDV = Right$(pInscricao, 1)
    Let peso = "98765432"
    
    For pos = 1 To 8
        Let soma = soma + (Val(Mid$(ieNum, pos, 1)) * Val(Mid$(peso, pos, 1)))
    Next pos
    
    Let resto = (soma Mod 11)
    
    If (resto < 2) Then
        Let ieDV = 0
    Else
        Let ieDV = (11 - resto)
    End If
    
    If (ieDV = origDV) Then
        ValidarSC = "Inscrição válida"
    Else
        ValidarSC = "Inscrição inválida, digito verificador = " & ieDV
    End If
End Property

Property Get ValidarGO() As String
    'Segundo instruções do Sintegra:
    'http://www.sintegra.gov.br/Cad_Estados/cad_GO.html
    Dim ieNum As String, origDV As String, peso As String, ieDV As String
    Dim pos As Integer, soma As Integer, resto As Integer
    
    Let ieNum = Left$(pInscricao, 8)
    Let origDV = Right$(pInscricao, 1)
    Let peso = "98765432"
    
    For pos = 1 To 8
        Let soma = soma + (Val(Mid$(ieNum, pos, 1)) * Val(Mid$(peso, pos, 1)))
    Next pos
    
    Let resto = (soma Mod 11)
    
    If (resto = 0) Then
        Let ieDV = 0
    ElseIf (resto = 1) Then
        If ((Val(ieNum) > 10103104) Or (Val(ieNum) < 10119998)) Then
            Let ieDV = 1
        Else
            Let ieDV = 0
        End If
    Else
        Let ieDV = (11 - resto)
    End If
    
    If (ieDV = origDV) Then
        ValidarGO = "Inscrição válida"
    Else
        ValidarGO = "Inscrição inválida, digito verificador = " & ieDV
    End If
End Property

Property Get ValidarAP() As String
    'Segundo instruções do Sintegra:
    'http://www.sintegra.gov.br/Cad_Estados/cad_AP.html
    Dim ieNum As String, origDV As String, peso As String, ieDV As String
    Dim pos As Integer, soma As Integer, resto As Integer, p As Integer, d As Integer
    
    Let ieNum = Left$(pInscricao, 8)
    Let origDV = Right$(pInscricao, 1)
    Let peso = "98765432"

    If (Val(ieNum) < Val("03017001")) Then
        Let p = 5
        Let d = 0
    ElseIf (Val(ieNum) < Val("03019023")) Then
        Let p = 9
        Let d = 1
    Else
        Let p = 0
        Let d = 0
    End If
    
    Let soma = p
    
    For pos = 1 To Len(ieNum)
        Let soma = soma + (Val(Mid$(ieNum, pos, 1)) * Val(Mid$(peso, pos, 1)))
    Next pos
    
    Let resto = (soma Mod 11)
    
    If (resto = 0) Then
        Let ieDV = d
    ElseIf (resto = 1) Then
        Let ieDV = 0
    Else
        Let ieDV = (11 - resto)
    End If
    
    Debug.Print "RESTO: " & resto, "P:" & p, "D:" & d, "DV:"; (11 - resto)
    
    If (ieDV = origDV) Then
        ValidarAP = "Inscrição válida"
    Else
        ValidarAP = "Inscrição inválida, digito verificador = " & ieDV
    End If
End Property

Property Get ValidarPR() As String
    'Segundo instruções do Sintegra:
    'http://www.sintegra.gov.br/Cad_Estados/cad_PR.html
    Dim ieNum As String, ieDV As String, peso As String, DV As String
    Dim pos As Integer, PosDV As Integer, soma As Integer, resto As Integer

    Let ieDV = Right$(pInscricao, 2)
    Let DV = ""
    
    For PosDV = 1 To Len(ieDV)
        
        If (PosDV = 2) Then
            Let peso = "432765432"
            Let ieNum = Left$(pInscricao, 9)
        Else
            Let peso = "32765432"
            Let ieNum = Left$(pInscricao, 8)
        End If
        
        Let soma = 0
        
        For pos = 1 To Len(peso)
            Let soma = soma + (Val(Mid$(ieNum, pos, 1)) * Val(Mid$(peso, pos, 1)))
        Next pos

        Let resto = 11 - (soma Mod 11)
        
        If resto > 9 Then Let resto = 0
        
        Let DV = DV & resto
        
    Next PosDV
    
    If (DV = ieDV) Then
        ValidarPR = "Inscrição válida"
    Else
        ValidarPR = "Inscrição inválida, digito verificador = " & DV
    End If
End Property

Property Get ValidarRS() As String
    'Segundo instruções do Sintegra:
    'http://www.sintegra.gov.br/Cad_Estados/cad_RS.html
    Dim ieNum As String, ieDV As String, peso As String, DV As String
    Dim pos As Integer, soma As Integer, resto As Integer
    
    Const nDV = 1
    Let ieNum = Left$(pInscricao, (Len(pInscricao) - nDV))
    Let ieDV = Right$(pInscricao, nDV)
    Let peso = "298765432"
    
    For pos = 1 To (Len(pInscricao) - nDV)
        Let soma = soma + (Val(Mid$(ieNum, pos, 1)) * Val(Mid$(peso, pos, 1)))
    Next pos
    
    Let resto = 11 - (soma Mod 11)
    If resto < 2 Then Let resto = 0
        
    Let DV = DV & resto
    
    If (DV = ieDV) Then
        ValidarRS = "Inscrição válida"
    Else
        ValidarRS = "Inscrição inválida, digito verificador = " & DV
    End If
    
End Property

Property Get ValidarMG() As String
    'Segundo instruções do Sintegra:
    'http://www.sintegra.gov.br/Cad_Estados/cad_MG.html
    
    Dim ieNum As String, ieDV As String, DV As String, somaStr As String
    Dim pos As Integer, soma As Integer, resto As Integer
    Dim peso As Variant
    Dim wf As WorksheetFunction
    
    Set wf = Application.WorksheetFunction
    
    Const nDV = 2
    Let ieNum = Left$(pInscricao, (Len(pInscricao) - nDV))
    Let ieDV = Right$(pInscricao, nDV)
    Let peso = "121212121212"
    Let somaStr = ""
    Let soma = 0
    Let DV = ""
    
    'Cálculo do primeiro dígito verificador
    Let ieNum = Left$(ieNum, 3) & "0" & Right$(ieNum, 8)
    
    For pos = 1 To Len(ieNum)
        Let somaStr = somaStr & (Val(Mid$(ieNum, pos, 1)) * Val(Mid$(peso, pos, 1)))
    Next pos
    
    For pos = 1 To Len(somaStr)
        Let soma = soma + Val(Mid$(somaStr, pos, 1))
    Next pos
    
    Let resto = wf.Ceiling(soma, 10) - soma
    Let DV = DV & resto
    
    'Cálculo do segundo dígito verificador
    Let ieNum = Left$(pInscricao, (Len(pInscricao) - nDV)) & DV
    Let peso = Split("3,2,11,10,9,8,7,6,5,4,3,2", ",")

    Let soma = 0
    For pos = 0 To UBound(peso)
        Let soma = soma + (Val(Mid$(ieNum, (pos + 1), 1)) * Val(peso(pos)))
    Next pos
    
    Let resto = 11 - (soma Mod 11)
    If ((soma Mod 11) < 2) Then Let resto = 0
    Let DV = DV & resto

    If (DV = ieDV) Then
        ValidarMG = "Inscrição válida"
    Else
        ValidarMG = "Inscrição inválida, digito verificador = " & DV
    End If
    
    Set wf = Nothing
    
End Property

Property Get ValidarSP() As String
    'Segundo instruções do Sintegra:
    'http://www.sintegra.gov.br/Cad_Estados/cad_SP.html
    
    Dim ieNum As String, ieDV As String, origDV As String
    Dim pos As Integer, soma As Integer, resto As Integer
    Dim peso As Variant

    'Inicializando os valores das variáveis
    Let ieNum = ""
    Let origDV = Mid$(pInscricao, 9, 1) & Mid$(pInscricao, 12, 1)
    Let peso = ""
    Let soma = 0
    Let ieDV = ""
    
    'Cálculo do primeiro dígito verificador posição 9
    Let ieNum = Left$(pInscricao, 8)
    Let peso = Split("1,3,4,5,6,7,8,10", ",")
    
    For pos = 0 To UBound(peso)
        Let soma = soma + (Val(Mid$(ieNum, (pos + 1), 1)) * Val(peso(pos)))
    Next pos
    
    Let resto = Right((soma Mod 11), 1)
    Let ieDV = ieDV & resto

    'Cálculo do segundo dígito verificador posição 12
    Let ieNum = Left$(pInscricao, 11)
    Let peso = Split("3,2,10,9,8,7,6,5,4,3,2", ",")
    Let soma = 0
    
    For pos = 0 To UBound(peso)
        Let soma = soma + (Val(Mid$(ieNum, (pos + 1), 1)) * Val(peso(pos)))
    Next pos
    
    Let resto = Right((soma Mod 11), 1)
    Let ieDV = ieDV & resto
    
    If (ieDV = origDV) Then
        ValidarSP = "Inscrição válida"
    Else
        ValidarSP = "Inscrição inválida, digito verificador = " & ieDV
    End If
    
End Property

Property Get ValidarRJ() As String
    'Segundo instruções do Sintegra:
    'http://www.sintegra.gov.br/Cad_Estados/cad_RJ.html
    
    Dim ieNum As String, ieDV As String, origDV As String
    Dim pos As Integer, soma As Integer, resto As Integer
    Dim peso As Variant

    'Inicializando os valores das variáveis
    Let ieNum = Left$(pInscricao, 7)
    Let origDV = Right$(pInscricao, 1)
    Let peso = "2765432"
    Let soma = 0
    Let ieDV = ""
 
    'Percorre as strings ieNum e peso aplicando a soma
    For pos = 1 To Len(peso)
        Let soma = soma + (Val(Mid$(ieNum, pos, 1)) * Val(Mid$(peso, pos, 1)))
    Next pos
    
    'Se o resto da soma divido por 11 for menor que 2 então será 0
    'Caso contrário será 11 menos o resto da divisão
    If ((soma Mod 11) < 2) Then
        Let ieDV = ieDV & 0
    Else
        Let ieDV = ieDV & (11 - (soma Mod 11))
    End If
    
    If (ieDV = origDV) Then
        ValidarRJ = "Inscrição válida"
    Else
        ValidarRJ = "Inscrição inválida, digito verificador = " & ieDV
    End If
    
End Property

Property Get ValidarDF() As String
    'Segundo instruções do Sintegra:
    'http://www.sintegra.gov.br/Cad_Estados/cad_DF.html
    Dim ieNum As String, ieDV As String, calcDV As String
    Dim pos As Integer, soma As Integer, resto As Integer
    Dim peso As Variant
    
    'Inicialização das variáveis
    Let ieNum = Left$(pInscricao, 11)
    Let ieDV = Right$(pInscricao, 2)
    Let peso = "43298765432"
    Let calcDV = ""
    Let soma = 0
    Let resto = 0
    
    'Cálculo do primeiro dígito verificador
    For pos = 1 To Len(ieNum)
        Let soma = soma + (Val(Mid$(ieNum, pos, 1)) * Val(Mid$(peso, pos, 1)))
    Next pos
    
    Let resto = (soma Mod 11)
    
    If (resto < 2) Then
        Let calcDV = 0
    Else
        Let calcDV = (11 - resto)
    End If
    
    'Cálculo do segundo dígito verificador
    Let ieNum = Left$(pInscricao, 11) & calcDV
    Let peso = "543298765432"

    Let soma = 0
    For pos = 1 To Len(ieNum)
        Let soma = soma + (Val(Mid$(ieNum, pos, 1)) * Val(Mid$(peso, pos, 1)))
    Next pos
    
    Let resto = (soma Mod 11)
    
    If (resto < 2) Then
        Let calcDV = calcDV & 0
    Else
        Let calcDV = calcDV & (11 - resto)
    End If

    If (calcDV = ieDV) Then
        ValidarDF = "Inscrição válida"
    Else
        ValidarDF = "Inscrição inválida, digito verificador = " & calcDV
    End If
End Property

'Property Get ValidarAM() As String
'End Property
'
'Property Get ValidarBA() As String
'End Property
'
'Property Get ValidarMT() As String
'End Property
'
'Property Get ValidarPE() As String
'End Property
'
'Property Get ValidarRN() As String
'End Property
'
'Property Get ValidarRO() As String
'End Property
'
'Property Get ValidarRR() As String
'End Property
'
'Property Get ValidarSE() As String
'End Property
'
'Property Get ValidarTO() As String
'End Property

