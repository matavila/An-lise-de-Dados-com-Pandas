#Primeiro passo: Importando as bibliotecas
import pandas as pd                         #Manipulação/Tratamento de dados             
import matplotlib.pyplot as plt             #Customização de gráficos

#Definindo a biblioteca Seaborn do plotlib
plt.style.use("seaborn")

#Fazendo a leitura do arquivo
dataframe= pd.read_excel('AdventureWorks.xlsx')

#Visualizando as Primeiras Linhas
print(dataframe.head(10))

#Verificando a quantiade de linhas e colunas
print(dataframe.shape)

#Verificando os tipos de dados
print(dataframe.dtypes)

#Verificando Venda total
Receita= dataframe['Valor Venda'].sum()
print(f"O valor da receita foi: R${Receita:0.2f}")

#verificando o Custo Total
dataframe["Custo"] = dataframe["Custo Unitário"].mul(dataframe['Quantidade'])
CustoTotal = dataframe["Custo"].sum()
print(f"Cuto total é de: R${CustoTotal:.01f}")

#Gerando a Receita total
dataframe["Receita"] = dataframe["Valor Venda"]-dataframe["Custo"]
ReceitaTotal = dataframe["Receita"].sum()
print(f"A receita total das vendas foi: R${ReceitaTotal:.03f}")

#Definindo tempo de entrega
dataframe["Tempo Entrega"] = dataframe["Data Envio"] - dataframe["Data Venda"] 
Tempo = dataframe["Tempo Entrega"].mean()
print(f'O valor média de tempo {Tempo}')

#Definindo a média de tempo de entrega para cada marca
dataframe["Tempo Entrega"]= dataframe['Tempo Entrega'].dt.days                      #Transformando as string 10 days em 10
print(dataframe['Tempo Entrega'].dtype)                                             #Verificando se foi transformado

#Fazendo a pre-selção (agrupamento) das marcas e tirando as médias
print(dataframe.groupby('Marca')["Tempo Entrega"].mean())

#Verificando se há valores vazios na base de dados
NunDadosVazios = dataframe.isnull().sum()
print(f"Quantidade de dados vazios: {NunDadosVazios}")

#Definindo o lucro com base no ano e na marca
dataframe['Ano'] = dataframe["Data Venda"].dt.year
pd.options.display.float_format = '{:20,.2f}'.format                                #Configurando o formato dos números flutuantes
print(dataframe.groupby(["Ano","Marca"])['Receita'].sum())
print("==================x==================")

#Resetando index e trasnformando as inforamações acima na forma de tabela
Lucro_Ano = dataframe.groupby(["Ano","Marca"])['Receita'].sum().reset_index()
print(Lucro_Ano)

#Verificando as quantidades vendidas por tipo de produto
Produto_Vendido = dataframe.groupby('Produto')['Quantidade'].sum().sort_values(ascending=False).reset_index()   #Trazendo do maior pro menor
print(Produto_Vendido)

#Gráfico total de produtos vendidos
Dados = dataframe.groupby('Produto')['Quantidade'].sum().plot.barh(title ="Total de produtos Vendidos")
plt.xlabel('Total')
plt.ylabel('Produto')
plt.show()

#Gráfico da Receita por ano
Dados2 = dataframe.groupby('Ano')["Receita"].sum().plot.bar(title ="Lucro por Ano")
plt.xlabel('Ano')
plt.ylabel('Total')
plt.show()

#Gráfico de receita selecionando somente um ano
Dados_2009 = dataframe[dataframe["Ano"]== 2019]
Dados_2009["Mes"] = dataframe["Data Venda"].dt.month

Dados3 = dataframe.groupby(Dados_2009['Mes'])['Receita'].sum().plot.bar(title="Lucro de 2019 por Mês")
plt.xlabel('Mês')
plt.ylabel('Total')
plt.show()

#Fazendo diversos gráficos
print(dataframe["Tempo Entrega"].describe())

#-> Boxplot
plt.boxplot(dataframe["Tempo Entrega"])
plt.show()

#-> Histograma
plt.hist(dataframe["Tempo Entrega"])
plt.show()

dataframe.to_csv('Df_Vendas_Novo.csv',index=False)