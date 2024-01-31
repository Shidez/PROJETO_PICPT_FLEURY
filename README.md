# PROJETO_POSTAL_PDF
Programa para facilitar a separação de resultados em pdf para impressão e envio ao cliente

Este projeto foi criado a partir da necessidade de facilitar a separação de resultados prontos para impressão e postagem aos clientes. 

Este projeto inclui alguns passos:

1. O preenchimento da planilha com os dados do cliente e de seu médico, incluindo e-mail.
2. A planilha contém codigo de VBA e a partir do número da ficha do cliente cria um e-mail automático para a liberação dos médicos responsaveis pelo exame, após o envio do e-mail coloca um OK na coluna "E-mail enviado" na planilha de controle.
3. A coluna POSTAL informa pelo código da unidade (três primeiros números de ficha) se aquela unidade tem acesso ao resultado pela internet ou não. Se não tiver acesso, a planilha preenche como "POSTAL" a coluna POSTAL da planilha.
4. Se for unidade que não tem acesso ao resultado via sistema, será necessário a impressão do resultado e envio via postal. 
5. O programa com python, lê a planilha "CONTROLE DE ENVIO 2023 EXAME_X" nos campos: Email enviado, Postal e Postal enviado para fazer uma cópia do resultado em pdf em outra pasta para facilitar a triagem e impressão dos resultados.
6. Se as colunas "Email enviado" e "Postal" estiverem preenchidas e a coluna de "Postal enviado", estiver vazia, o programa separa o resultado do cliente com o número da ficha (ex.: 123.pdf) e coloca um "OK" na coluna de Postal enviado para não separar o resultado novamente.

Com isto, otimizamos o tempo e garantimos a separação correta dos pdfs, amenizando erros humanos. 


