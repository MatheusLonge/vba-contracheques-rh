Sub CopiarLinhasAposData()
    Dim caminhoBase As String
    Dim dataInicio As String
    Dim nomeFuncionario As String
    Dim linhasCopiadas As String
    Dim copiando As Boolean
    Dim linha As String
    Dim ws As Worksheet
    Dim i As Integer
    Dim ano As Integer
    Dim meses As Variant
    Dim linhaAtual As Long
    Dim naoEncontrados As String
    Dim blocoEncontrado As Boolean
    Dim primeiraLinha As String
    Dim segundaLinha As String
    Dim contadorLinhas As Integer
    
    ' Caminho base onde as pastas de anos estão localizadas
    caminhoBase = "G:\ESPELHO DA FOLHA SDE.SUDIC.CIS\"
    dataInicio = ThisWorkbook.Sheets("PLAN2").Range("B2").Value
    nomeFuncionario = ThisWorkbook.Sheets("PLAN2").Range("A2").Value
    Set ws = ThisWorkbook.Sheets("Plan10")
    
    ' Array com os nomes dos meses
    meses = Array("JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ")
    
    ' Inicializar a linha atual da planilha e a string de meses/anos não encontrados
    linhaAtual = 1
    naoEncontrados = ""
    
    ' Percorrer os anos de 1994 a 2018
    For ano = 1994 To 2018
        ' Adicionar o cabeçalho separador para o novo ano
        ws.Cells(linhaAtual, 1).Value = "CONTRACHEQUES " & ano
        linhaAtual = linhaAtual + 1
        
        ' Percorrer cada mês no array
        For i = LBound(meses) To UBound(meses)
            linhasCopiadas = ""
            copiando = False
            blocoEncontrado = False
            primeiraLinha = ""
            segundaLinha = ""
            contadorLinhas = 0
            
            ' Construir o caminho completo do arquivo
            Dim caminhoArquivo As String
            caminhoArquivo = caminhoBase & ano & "\" & meses(i) & ano
            
            ' Verificar se o arquivo existe antes de tentar abri-lo
            If Dir(caminhoArquivo) <> "" Then
                ' Abrir o arquivo para leitura
                Open caminhoArquivo For Input As #1
                
                ' Ler as duas primeiras linhas
                If Not EOF(1) Then
                    Line Input #1, primeiraLinha
                    contadorLinhas = contadorLinhas + 1
                End If
                If Not EOF(1) Then
                    Line Input #1, segundaLinha
                    contadorLinhas = contadorLinhas + 1
                End If
                
                ' Percorrer o arquivo linha por linha
                Do Until EOF(1)
                    Line Input #1, linha
                    
                    ' Verificar se estamos copiando
                    If copiando Then
                        ' Verificar se encontramos o final do bloco
                        If InStr(1, linha, "MG.CONSIG.") > 0 Or InStr(1, linha, "OUTROS AFASTAMENTOS") > 0 Or InStr(1, linha, "EXERCICIO DE MANDATO ELETIV") > 0 Then
                            copiando = False
                            
                            ' Colar o bloco apenas se `blocoEncontrado` for verdadeiro
                            If blocoEncontrado Then
                                ' Colar as duas primeiras linhas como cabeçalho
                                ws.Cells(linhaAtual, 1).Value = primeiraLinha
                                linhaAtual = linhaAtual + 1
                                ws.Cells(linhaAtual, 1).Value = segundaLinha
                                linhaAtual = linhaAtual + 1
                                
                                ' Colar o bloco de linhas copiadas
                                ws.Cells(linhaAtual, 1).Value = linhasCopiadas
                                linhaAtual = linhaAtual + 1
                                
                                ' Limpar as variáveis do bloco
                                linhasCopiadas = ""
                                blocoEncontrado = False
                            End If
                        End If
                        
                        ' Verificar se encontramos o nome do funcionário correto
                        If InStr(1, linha, nomeFuncionario) > 0 Then
                            blocoEncontrado = True
                        End If
                        
                        ' Verificar se encontramos um novo funcionário ou matrícula diferente
                        If InStr(1, linha, "0701") > 0 And Not InStr(1, linha, nomeFuncionario) > 0 Then
                            copiando = False
                            blocoEncontrado = False
                            linhasCopiadas = ""
                        End If
                        
                        ' Continuar copiando as linhas se ainda estamos no bloco do funcionário correto
                        If blocoEncontrado Then
                            linhasCopiadas = linhasCopiadas & linha & vbCrLf
                        End If
                    End If
                    
                    ' Verificar se encontramos a data de início
                    If InStr(1, linha, dataInicio) > 0 Then
                        copiando = True
                        linhasCopiadas = linhasCopiadas & linha & vbCrLf
                    End If
                Loop
                
                ' Fechar o arquivo
                Close #1
                
                ' Caso o arquivo termine sem encontrar "MG.CONSIG." ou "OUTROS AFASTAMENTOS", copiar bloco pendente
                If blocoEncontrado Then
                    ' Colar as duas primeiras linhas como cabeçalho
                    ws.Cells(linhaAtual, 1).Value = primeiraLinha
                    linhaAtual = linhaAtual + 1
                    ws.Cells(linhaAtual, 1).Value = segundaLinha
                    linhaAtual = linhaAtual + 1
                    
                    ' Colar o bloco de linhas copiadas
                    ws.Cells(linhaAtual, 1).Value = linhasCopiadas
                    linhaAtual = linhaAtual + 1
                    
                    ' Limpar as variáveis do bloco
                    linhasCopiadas = ""
                    blocoEncontrado = False
                End If
                
            Else
                ' Adicionar à lista de não encontrados
                naoEncontrados = naoEncontrados & meses(i) & " " & ano & vbCrLf
            End If
        Next i
    Next ano
    
    ' Exibir uma mensagem informando que o processo foi concluído
    MsgBox "Linhas copiadas foram coladas na planilha."
    
    ' Exibir uma mensagem informando os meses/anos não encontrados
    If naoEncontrados <> "" Then
        MsgBox "Os seguintes meses/anos não foram encontrados:" & vbCrLf & naoEncontrados
    End If
End Sub
