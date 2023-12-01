# SistemaOrcamentario

## Descrição do Projeto
Sistema realizado para agilizar a criação de orçamentos, sendo suas funções:
- Cria orçamentos
- Edita orçamentos
- Insere, edita e exclui dados em banco Access

## Tecnologias utilizadas
- VBA
- Access
  
##  Descrição Detalhada do Projeto

O descrição do projeto apresentada será realizada em duas etapa, sendo elas as principais que componhem ele, sendo: Planilha Excel e suas guias e Formulário do projeto. 

### Sobre as Guias
O sistema possui 3 guias principais, sendo elas:

### 1.Menu
Responsável por direcionar o usuário através do sistema por meio dos botões.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/e4ada0e0-5917-4b2e-be67-9ed553583184)




### 2. Criação de Orçamentos

Possibilita a adição, edição e exclusão de linhas de um orçamento.

- Ao clicar no botão de adição é possível iniciar um novo orçamento partindo do zero.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/8da7cd4a-e574-43a9-84d1-cd842260f37c)
- Ao clicar em editar é possível adicionar um novo serviço dentro de um orçamento já criado, acessando a lista e "flegando" o serviço.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/bd561ef1-a266-449e-85dd-459b4b792f5b)

- Ao clicar em excluir, é possível excluir um serviço em que tenha acontecido a marcação dentro da célula marcada em verda na imagem abaixo.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/b5733d0a-23db-464b-ba23-61797d168635)



### 3. Custos

A guias custos possibilida verificar os custos em que cada serviço irá gerar, como também adicionar ou excluir o valor da mão de obra utilizada.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/c7b8856b-e8da-436b-81cc-c6fbc87e0971)

##  Formulário do Projeto

Na guia principal, o botão de "Adicionar Insumo" irá realizar a execução da abertura do formulário.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/fdeff28e-46c8-415d-8dec-060df5857117)

O formulário conta com as 3 funções básicas: adição, edição e exclusão.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/dad9ed0b-c5ca-41a5-b9d4-c6009d611b7f)



### Adição
Permite que um novo serviço seja criado, de acordo com a regra de negócio para esse projeto em particular, é possível escolher 3 tipos diferentes de serviços.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/212fffb2-cdb0-4647-ab41-9ae86b4d5ff4)

Ao escolher um deles, serão preenchidos 3 níveis + 1 sendo o insumo. </br>
Os campos que podem ser utilizados são uma seleção em lista (dados são puxados direto do dB) ou um novo serviço pode ser digitado no campo de tipo texto. </br>
O serviço irá prosseguir ao clicar no botão "avançar" apenas se não houver nenhum serviço similar ao preenchido nos campos.

![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/9077049e-0a21-412f-87eb-d650e4f6f4a3)

Ao clicar em avançar, uma nova área é habilitada para a inserção dos dados do serviço.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/d64bcf8a-cf57-485e-87d1-b55af81fed58)

Se todos os campos forem preenchidos, o serviço poderá ser salvo.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/79614135-f0ce-4952-9590-d963e2d1f36b)


### Edição

Para edição, a separação é realizada por tipo de serviço (em verde), cada qual com suas opções de edição (em vermelho). </br>
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/6839dc2e-8d27-4cc3-86fd-1f908ddf3d87) </br> </br>

Para edição de denominação de nível 1,2 e 3. Os 3 botões abaixo são acessados de acordo com a necessidade de edição. </br>
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/512415f9-573e-4038-89dc-0cd587265a55) </br> </br>
Em vermelho está sinalizado a lista com a denominação a ser modificada, em verde o campo texto pode ser utilizado para digitar a nova denominação para o nível. </br>
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/e85fa19f-514d-4d2d-bb34-7e742bd63ef8) </br> </br>

Para insumo, é possível editar a edição para os dados associados a ele pelo botão "Editar Insumo | Unidade | Custo Insumo". </br>
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/8ae84814-03cb-415d-863a-9b5c1cd771e7) </br> </br>
Nesta guia, 4 informações associadas ao insumo podem ser alteradas, tal como print abaixo. </br> </br>
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/4b56b9ce-1860-4743-b94f-8bf7645c11fd) </br> </br>

Para edição do "Preço de mão de obra" e "Preço de venda sugerido" o botão abaixo é acessado.

![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/7f2ead1b-98d9-46d8-a9d4-690bd5d285b6) </br> </br>

Para editar ambos os valores, é necessário que uma combinação de níveis (que forma um serviço existente) seja preenchida. 
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/93173718-c373-45ff-9573-918ef136c158) </br> </br>

Após o preenchimento do primeiro nível, os níveis subsequentes serão preenchidos utilizando como busca o Id dos níveis anteriores no banco de dados. Logo, sempre trará para a lista dados que já tenham registro de serviço no banco de dados, por meio desta busca que atualiza cada lista subsequente com base na(s) selecionadas anteriormente.</br>

Após a seleção dos níveis, será habilitado a edição para os valores de preço de venda sugerido e custo de mão de obra.

![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/595cdae7-5269-4a4b-afcd-d7c2e83b3051) </br> </br>

** Rendimento **

![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/98c7b311-13d0-42f7-bd27-f56b0c5768fd)


### Exclusão

Para cada tipo de serviço e para os insumo (em vermelho), estão habilitados algumas opções para edição de acordo com a regra de negócio para aquele determinado tipo de serviço.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/53b69318-a7cd-4d50-8379-662438d08f89)

Para o nível 1, 2, 3 e insumo, apenas um destes níveis e insumo poderá ser excluído se não houverem serviços associados a ele. </br> 
Caso houverem itens associado ao nível ou insumo selecionado para exclusão, uma lista será criada com todos os serviços que possuem esse item para que seja usado no botão de "exclusão de estrutura". </br>
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/ef104bdd-5162-493a-88ca-c21c425e63a4)
Caso não exista nenhum serviço com o item associado, será liberado a exclusão. </br>

Caso necessário a exclusão da estrutra, como mencionado anteriormente, será necessário entrar no botão de exclusão de estrutura para aquele tipo de serviço em que se deseja realizar a exclusão do nível.

No botão de exclusão de estrutura, estarão disponíveis 3 listas que trazem dados direto do banco de dados.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/5fcfcc4a-78dd-4ad8-97e2-9b1e092732a9)

Ao inserir uma estrutura existente, é identificado a existencia e possibilidade de exclusão.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/06903a60-f6a9-4ab3-a709-b72a1703df97)
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/5eec97cd-dd94-488f-9b3e-950125d412a1)

Ao finalizar uma exclusão, a lista é atualizada e o número de estruturas com aquele nível é atualizada.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/49425c51-1701-40ba-847b-b0c77f1b035f)
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/081839b9-90c9-444e-a04e-8800e1f1153d)

Estrutura similar é utilizada para outros tipos de serviço e para exclusão de insumo.
