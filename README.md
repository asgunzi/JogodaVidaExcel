# JogodaVidaExcel
Implementação do jogo da vida em Excel VBA


O matemático John Conway faleceu na semana passada, vítima do Corona vírus.

https://guiadoestudante.abril.com.br/estudo/conheca-john-conway-o-matematico-que-criou-o-jogo-da-vida/

Ele foi o criador do “Jogo da Vida”, o primeiro exemplo de autômato celular. É bastante interessante e lúdico.

Em homenagem a Conway, é mais ou menos simples fazer uma versão em Excel.

Imagine um tabuleiro, com pontos aleatórios.
![](https://forgottenmathhome.files.wordpress.com/2020/04/gamelife01.jpg)


O jogo faz a seguinte análise:

    Qualquer célula viva com menos de dois vizinhos vivos morre de solidão.
    Qualquer célula viva com mais de três vizinhos vivos morre de superpopulação.
    Qualquer célula morta com exatamente três vizinhos vivos se torna uma célula viva.
    Qualquer célula viva com dois ou três vizinhos vivos continua no mesmo estado para a próxima geração.

Rodando várias iterações, começam a surgir alguns padrões.

![](https://forgottenmathhome.files.wordpress.com/2020/04/gravar-_2.gif?w=1024)


O código gera uma matriz aleatória de 0 e 1. Cola na planilha.

Pinta de verde e branco, com formatação condicional.

A seguir, aplica as regras do jogo da vida.

Repita o procedimento acima diversas vezes.

Vamos analisar apenas um trecho do código, as regras do jogo da vida.

Existe uma variável chamada somaVizinhos que soma quantos vizinhos a célula tem.

A célula em referência é arrVal(i, j). Se ela for igual a 1, está viva, se 0, não.

‘Qualquer célula viva com menos de dois vizinhos vivos morre de solidão.

            If arrVal(i, j) = 1 And somaVizinhos < 2 Then

                arrValUpdate(i, j) = 0

            End If

            ‘Qualquer célula viva com mais de três vizinhos vivos morre de superpopulação.

            If arrVal(i, j) = 1 And somaVizinhos > 3 Then

                arrValUpdate(i, j) = 0

            End If

            ‘Qualquer célula morta com exatamente três vizinhos vivos se torna uma célula viva.

            If arrVal(i, j) = 0 And somaVizinhos = 3 Then

                arrValUpdate(i, j) = 1

            End If

            ‘Qualquer célula viva com dois ou três vizinhos vivos continua no mesmo estado para a próxima geração.

            If arrVal(i, j) = 1 And (somaVizinhos = 2 Or somaVizinhos = 3) Then

                arrValUpdate(i, j) = 1

            End If


Lição de casa:  modificar a terceira regra para:

‘Qualquer célula morta com exatamente TRÊS OU QUATRO vizinhos vivos se torna uma célula viva’ e analisar os resultados.

![](https://forgottenmathhome.files.wordpress.com/2020/04/gamelife02.jpg?w=932)


Bom divertimento!
