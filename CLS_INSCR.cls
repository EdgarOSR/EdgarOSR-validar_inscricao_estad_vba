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

Private Sub Class_Initialize()
    pInscricao = ""
End Sub

Property Get Inscricao() As String
    Inscricao = pInscricao
End Property

Property Let Inscricao(ByVal ValorTextbox As String)
    pInscricao = Replace(Replace(ValorTextbox, ".", ""), "-", "")
End Property

Property Get ValidarSC() As String

    Dim IeNum As String, IeDV As String, Peso As String, DV As String
    Dim Pos As Integer, Soma As Integer, Resto As Integer
    
    Let IeNum = Left$(pInscricao, 8)
    Let IeDV = Right$(pInscricao, 1)
    Let Peso = "98765432"
    
    For Pos = 1 To 8
        Let Soma = Soma + (Val(Mid$(IeNum, Pos, 1)) * Val(Mid$(Peso, Pos, 1)))
    Next Pos
    
    Let Resto = 11 - (Soma Mod 11)
    
    If Resto < 2 Then Let Resto = 0
        
    Let DV = DV & Resto
    
    If (DV = IeDV) Then
        ValidarSC = "Inscrição válida"
    Else
        ValidarSC = "Inscrição inválida, digito verificador = " & DV
    End If
    
End Property

Property Get ValidarPR() As String

    Dim IeNum As String, IeDV As String, Peso As String, DV As String
    Dim Pos As Integer, PosDV As Integer, Soma As Integer, Resto As Integer

    Let IeDV = Right$(pInscricao, 2)
    Let DV = ""
    
    For PosDV = 1 To 2
        
        If (PosDV = 2) Then
            Let Peso = "432765432"
            Let IeNum = Left$(pInscricao, 9)
        Else
            Let Peso = "32765432"
            Let IeNum = Left$(pInscricao, 8)
        End If
        
        Let Soma = 0
        
        For Pos = 1 To Len(Peso)
            Let Soma = Soma + (Val(Mid$(IeNum, Pos, 1)) * Val(Mid$(Peso, Pos, 1)))
        Next Pos

        Let Resto = 11 - (Soma Mod 11)
        
        If Resto > 9 Then Let Resto = 0
        
        Let DV = DV & Resto
        
    Next PosDV
    
    If (DV = IeDV) Then
        ValidarPR = "Inscrição válida"
    Else
        ValidarPR = "Inscrição inválida, digito verificador = " & DV
    End If
    
End Property
