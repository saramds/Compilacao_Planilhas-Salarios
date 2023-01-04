Attribute VB_Name = "M�dulo2"
Function salariocomimposto(qtd_normal As Double, qtd_extra As Double, preco_normal As Double, preco_extra As Double) As Double

'Cria f�rmula para calcular os sal�rios dos funcion�rios

salario = qtd_normal * preco_normal + qtd_extra * preco_extra

'Condi��es de acr�scimo do imposto

If salario <= 12000 Then

    salariocomimposto = salario

ElseIf salario <= 18000 Then

    salariocomimposto = salario * 1.1

Else

    salariocomimposto = salario * 1.125

End If

End Function
Sub compilar_salarios()

'Declara as vari�veis

Dim valor_normal As Double

Dim valor_extra As Double

'Seleciona a aba principal

Sheets("Exemplo Funcion�rios").Activate

'Descobre os valores da hora normal e da hora extra

valor_normal = Range("H6").Value

valor_extra = Range("H7").Value

For Each aba In ThisWorkbook.Sheets

'Faz o c�lculo do sal�rio para as abas dos setores da empresa

    If aba.name <> "Exemplo Funcion�rios" Then
    
        aba.Activate
        
        linha = 2
        
        Do Until Cells(linha, 1).Value = ""
        
            Cells(linha, 4).Value = salariocomimposto(Cells(linha, 2).Value, Cells(linha, 3).Value, valor_normal, valor_extra)
            
            linha = linha + 1
        
        Loop
    
    End If

Next

Sheets("Exemplo Funcion�rios").Activate

'Volta para a aba principal

End Sub

Sub limpar_abas()

For Each aba In ThisWorkbook.Sheets

'Limpas as informa��es antigas das abas dos setores da empresa

    If aba.name <> "Exemplo Funcion�rios" Then
    
        aba.Activate
        
        linha = 2
        
        Do Until Cells(linha, 1).Value = ""
        
            Cells(linha, 4).ClearContents
            
            linha = linha + 1
        
        Loop
    
    End If

Next

Sheets("Exemplo Funcion�rios").Activate

'Volta para a aba principal

End Sub

