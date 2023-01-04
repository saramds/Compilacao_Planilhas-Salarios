Attribute VB_Name = "Módulo2"
Function salariocomimposto(qtd_normal As Double, qtd_extra As Double, preco_normal As Double, preco_extra As Double) As Double

salario = qtd_normal * preco_normal + qtd_extra * preco_extra

If salario <= 12000 Then

    salariocomimposto = salario

ElseIf salario <= 18000 Then

    salariocomimposto = salario * 1.1

Else

    salariocomimposto = salario * 1.125

End If

End Function
Sub compilar_salarios()

Dim valor_normal As Double

Dim valor_extra As Double

Sheets("Exemplo Funcionários").Activate

valor_normal = Range("H6").Value

valor_extra = Range("H7").Value

For Each aba In ThisWorkbook.Sheets

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

End Sub

Sub limpar_conteudos()

For Each aba In ThisWorkbook.Sheets

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

End Sub

