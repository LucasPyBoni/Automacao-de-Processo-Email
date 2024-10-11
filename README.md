## Sobre o Projeto:

Automação de arquivo que lê e mescla tabelas dimensão e fato do excel, criar pastas de cada loja e as envia por email para cada gerente dessas respectivas lojas.

### Principais bibliotecas:

Pandas, Pathlib, win32com e os

### Explicação da linha de raciocínio:

O programa lê e duas tabelas dimensão e uma tabela fato através do pandas.read. A tabela de loja mescla com a base de vendas, criando uma base de vendas para cada loja. A automação cria uma pasta, e posteriormente uma tabela p/ cada loja com nome e dia no titulo personalizados contendo meta do dia e do ano. O email para cada gerente acontece devido a biblioteca win32com e os filtros inteligentes por meio do pandas. Além disso, a tabela tem informações normais e também em formato html. E por fim é enviado um email especial para diretoria com os resultados das lojas agrupadas por ranking por meio do group by.
