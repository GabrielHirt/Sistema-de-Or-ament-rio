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
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/d6732463-0c1d-4cdd-acc8-105d2169e13d)


### Adição
Permite que um novo serviço seja criado, de acordo com a regra de negócio para esse projeto em particular, é possível escolher 3 tipos diferentes de serviços.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/212fffb2-cdb0-4647-ab41-9ae86b4d5ff4)

Ao escolher um deles, serão preenchidos 3 níveis + 1 sendo o insumo.
Os campos que podem ser utilizados são uma seleção em lista (dados são puxados direto do dB) ou um novo serviço pode ser digitado no campo de tipo texto.
O serviço irá prosseguir ao clicar no botão "avançar" apenas se não houver nenhum serviço similar ao preenchido nos campos
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/9077049e-0a21-412f-87eb-d650e4f6f4a3)

Ao clicar em avançar, uma nova área é habilitada para a inserção dos dados do serviço.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/d64bcf8a-cf57-485e-87d1-b55af81fed58)

Se todos os campos forem preenchidos, o serviço poderá ser salvo.
![image](https://github.com/GabrielHirt/SistemaOrcamentario/assets/98654562/79614135-f0ce-4952-9590-d963e2d1f36b)


### Edição

### Exclusão
