Attribute VB_Name = "Módulo2"
Function salariocomimposto(qtd_normal As Double, qtd_extra As Double, preco_normal As Double, preco_extra As Double) As Double

'Cria fórmula para calcular os salários dos funcionários

salario = qtd_normal * preco_normal + qtd_extra * preco_extra

'Condições de acréscimo do imposto

If salario <= 12000 Then

    salariocomimposto = salario

ElseIf salario <= 18000 Then

    salariocomimposto = salario * 1.1

Else

    salariocomimposto = salario * 1.125

End If

End Function
Sub compilar_salarios()

'Declara as variáveis

Dim valor_normal As Double

Dim valor_extra As Double

'Seleciona a aba principal

Sheets("Exemplo Funcionários").Activate

'Descobre os valores da hora normal e da hora extra

valor_normal = Range("H6").Value

valor_extra = Range("H7").Value

For Each aba In ThisWorkbook.Sheets

'Faz o cálculo do salário para as abas dos setores da empresa

    If aba.name <> "Exemplo Funcionários" Then
    
        aba.Activate
        
        linha = 2
        
        Do Until Cells(linha, 1).Value = ""
        
            Cells(linha, 4).Value = salariocomimposto(Cells(linha, 2).Value, Cells(linha, 3).Value, valor_normal, valor_extra)
            
            linha = linha + 1
        
        Loop
    
    End If

Next

Sheets("Exemplo Funcionários").Activate

'Volta para a aba principal

End Sub

Sub limpar_abas()

For Each aba In ThisWorkbook.Sheets

'Limpas as informações antigas das abas dos setores da empresa

    If aba.name <> "Exemplo Funcionários" Then
    
        aba.Activate
        
        linha = 2
        
        Do Until Cells(linha, 1).Value = ""
        
            Cells(linha, 4).ClearContents
            
            linha = linha + 1
        
        Loop
    
    End If

Next

Sheets("Exemplo Funcionários").Activate

'Volta para a aba principal

End Sub

