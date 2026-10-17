Desafio proposto:

Descrição do Desafio
Este projeto tem como objetivo criar uma ferramenta no Excel que ajude a organizar e reunir informações essenciais para a declaração de imposto de renda. A proposta é construir um agregador de dados no qual o usuário possa controlar suas entradas de maneira eficiente e validada, com menus de navegação, validações automáticas e funcionalidades extras, como links rápidos. A solução será completamente construída no Excel, com recursos que tornam a ferramenta robusta, mas com uma interface amigável e prática.

Estrutura do agregador:

1. Informações Pessoais 
2. Dados Bancários
3. Notas

Etapas:

1. Criação de menu de navegação

   Todas abas possuem o menu de navegação onde é possível clicar para ser direcionado para a página desejada. Conexão realizada pela função link no objeto.

2. Aba Titular

   Formatação de campos de preenchimento de acordo com necessidades de cada informação, sendo utilizados os seguintes tipos: texto, data, validação de dados por lista, CEP e customizadas. Formatações customizadas para os campos de CPF, título de eleitor, telefone e celular.

3. Aba informes

   Criação de menu para informar qual o banco, valor ao final do ano e anexo do Informe de Rendimentos. Criação de lista com validação de dados para preenchimento do campo banco.
   Nessa aba foi realizada a maior alteração em comparação ao modelo ensinado em aula. Adicionei um botão com macro (ModAdicionarBanco) para realizar adição de novos Informes de Rendimento, não limitando aos três bancos propostos pelo modelo.

4. Aba Notas
   Criação de tabela com para preenchimento de informações de recebimentos, aplicação principalmente de formatação: data, validação de dados e moeda.

5. Navegação e proteção da planilha
   Para além da navegação pelo menu também foi adicionada navegação com possibilidade de avançar e retornar entre as abas. Por fim, foi feita a seleção de células que necessitam de preenchimento para desbloqueá-las e permitir edição após proteção da planilha. Também foram feito ajustes para remover headings, gridlines e formula bar na visualização, além de ocultar o menu de abas.
