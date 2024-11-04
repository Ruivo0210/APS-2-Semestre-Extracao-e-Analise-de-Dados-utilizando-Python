# Aqui, estou importando os módulos necessários para o funcionamento do código,sendo eles: "matplotlib", reposnsável por gerar os gráficos em imagem; "pandas", responsável pela
# interação entre o python e as planilhas; e "openpyxl", responsável por gerar as novas planilhas baseadas nos dados analisados. Caso não possua os módulos, eles podem ser
# instalados com o comando "pip3 install".
import matplotlib.pyplot as plt
import pandas as pd
import openpyxl

# Aqui, estou definindo qual é a planilha para o módulo "pandas" poder ler, e logo em seguida definindo uma variável à aquela planilha, no formato de DataFrame.
planilha = 'Planilhas Censo 2021/Censo 2021 - Mesorregião de Ribeirão Preto.xlsx'
df_planilha = pd.read_excel(planilha)

# Como nem todas as escolas possuem dados informados, fiz com que o python me informasse a quantidade total de escolas que foram analisadas, a quantidade de escolas que não
# possuem dados informados (fiz com que o módulo "pandas" me informasse as células da planilha com dados vazios), e logo em seguida calculei a quantidade de escolas que possuem dados informados.
qntescolas = df_planilha[df_planilha.columns[0]].count()
qntescolasna = df_planilha.IN_TRATAMENTO_LIXO_INEXISTENTE.isnull().sum()
qntescolascenso = (qntescolas) - (qntescolasna)

# Aqui, estou imprimindo ao usuário os dados calculados acima.
print("Quantidade total de escolas na mesorregião de Ribeirão Preto: ", qntescolas)
print("Quantidade de escolas na mesorregião de Ribeirão Preto que possuem dados informados: ", qntescolascenso)
print("Quantidade de escolas na mesorregião de Ribeirão Preto que não possuem dados informados: ", qntescolasna)

# Como os dados calculados acima não são informados na planilha principal, fiz com que o módulo "openpyxl" criasse uma nova planilha baseada nos dados analisados.
df_escolascenso = pd.DataFrame([[qntescolas, qntescolascenso, qntescolasna]], index = ['Quantidade'], columns = ['Total de escolas', 'Dados informados', 'Dados não informados'])
df_escolascenso.to_excel('Novas planilhas geradas/PlanilhaTotaisCenso.xlsx', sheet_name = 'TotaisCenso')
print("Uma nova planilha com a quantidade total das escolas com dados informados foi gerada!")

# Agora que calculamos a quantidade total de escolas, escolas que possuem e que não possuem dados informados, montei uma lista baseada nessas variáveis, e fiz com que o módulo
# matplotlib analise e imprima um gráfico de barras para o usuário, baseado nos dados já analisados. O gráfico é salvo em um arquivo .png logo em seguida. Como o módulo matplotlib
# não possui nenhuma forma nativa de exibir a quantidade dos dados em cima de cada barra, usei uma estrutura de repetição "for" para mostrar os dados em cima das barras.
lista_escolas = [qntescolas, qntescolascenso, qntescolasna]
xlista = [x for x in range(len(lista_escolas))]
for i, v in enumerate(lista_escolas):
    plt.text(xlista[i] - 0.25, v + 0.01, str(v))
plt.bar(xlista, lista_escolas)
plt.title("Gráfico das escolas que possuem dados informados:")
plt.xlabel("Estado dos dados")
plt.ylabel("Quantidade de escolas")
plt.xticks([0, 1, 2], ['Total', 'Informado', 'Não informado'])
plt.savefig('Imagens geradas dos gráficos/Gráfico de escolas que possuem dados informados.png')
plt.show()

# A partir daqui, começamos a trabalhar com as cinco colunas correspondentes aos dados que quero analisar. Criei duas variáveis, em uma fiz com que o python calculasse a quantidade
# de dados que possuem em cada uma das possíveis representações na planilha ("Realizam " e "Não realizam"), e na outra pedi que ele calculasse a porcentagem entre esses dados. Em
# seguida, imprimi esses dados ao usuário, tive que usar o parâmetro ".to_string()" para evitar que o script mostrasse informações da variável (caso não esteja em string, ele exibe
# a tipagem da variável junto com os dados, informação desnecessária durante a análise), e pedi para que o módulo "matplotlib" imprimisse um gráfico em formato 
# de pizza baseado nessas informações.
qntescolaslixo = df_planilha['IN_TRATAMENTO_LIXO_INEXISTENTE'].value_counts()
pctescolaslixo = df_planilha['IN_TRATAMENTO_LIXO_INEXISTENTE'].value_counts(normalize = True).mul(100).apply(lambda x: "{:,.2f}".format(x))+'%'
print ("Quantidade de escolas que não realizam o tratamento do lixo: \n", qntescolaslixo.to_string())
print ("Porcentagem de escolas que não realizam o tratamento do lixo: \n", pctescolaslixo.to_string())
plt.pie(qntescolaslixo, autopct = lambda p : '{:.2f}%  ({:.0f})'.format(p, p * sum(qntescolaslixo)/100))
plt.title("Gráfico das escolas que não \nrealizam o tratamento do lixo:")
plt.legend(["Não realizam", "Realizam"], loc = 'lower left')
plt.savefig('Imagens geradas dos gráficos/Gráfico de escolas que não realizam o tratamento do lixo.png')
plt.show()

qntescolasref = df_planilha['IN_REFEITORIO'].value_counts()
pctescolasref = df_planilha['IN_REFEITORIO'].value_counts(normalize = True).mul(100).apply(lambda x: "{:,.2f}".format(x))+'%'
print ("Quantidade de escolas que possuem um refeitório: \n", qntescolasref.to_string())
print ("Porcentagem de escolas que possuem um refeitório: \n", pctescolasref.to_string())
plt.pie(qntescolasref, autopct = lambda p : '{:.2f}%  ({:.0f})'.format(p, p * sum(qntescolasref)/100))
plt.title("Gráfico das escolas que possuem um refeitório:")
plt.legend(["Possuem", "Não Possuem"], loc = 'lower left')
plt.savefig('Imagens geradas dos gráficos/Gráfico de escolas que possuem um refeitório.png')
plt.show()

qntescolaslab = df_planilha['IN_LABORATORIO_INFORMATICA'].value_counts()
pctescolaslab = df_planilha['IN_LABORATORIO_INFORMATICA'].value_counts(normalize = True).mul(100).apply(lambda x: "{:,.2f}".format(x))+'%'
print ("Quantidade de escolas que possuem um laboratório de informática: \n", qntescolaslab.to_string())
print ("Porcentagem de escolas que possuem um laboratório de informática: \n", pctescolaslab.to_string())
plt.pie(qntescolaslab, autopct = lambda p : '{:.2f}%  ({:.0f})'.format(p, p * sum(qntescolaslab)/100))
plt.title("Gráfico das escolas que possuem\n um laboratório de informática:")
plt.legend(["Não possuem", "Possuem"], loc = 'lower left')
plt.savefig('Imagens geradas dos gráficos/Gráfico de escolas que possuem um laboratório de informática.png')
plt.show()

qntescolasbanpne = df_planilha['IN_BANHEIRO_PNE'].value_counts()
pctescolasbanpne = df_planilha['IN_BANHEIRO_PNE'].value_counts(normalize = True).mul(100).apply(lambda x: "{:,.2f}".format(x))+'%'
print ("Quantidade de escolas que possuem banheiros acessíveis à pessoas com necessidades especiais: \n", qntescolasbanpne.to_string())
print ("Porcentagem de escolas que possuem banheiros acessíveis à pessoas com necessidades especiais: \n", pctescolasbanpne.to_string())
plt.pie(qntescolasbanpne, autopct = lambda p : '{:.2f}%  ({:.0f})'.format(p, p * sum(qntescolasbanpne)/100))
plt.title("Gráfico das escolas que possuem banheiros\n acessíveis à pessoas com necessidades especiais:")
plt.legend(["Possuem", "Não possuem"], loc = 'lower left')
plt.savefig('Imagens geradas dos gráficos/Gráfico de escolas que possuem banheiros acessíveis à pessoas com necessidades especiais.png')
plt.show()

qntescolaspne = df_planilha['IN_ACESSIBILIDADE_INEXISTENTE'].value_counts()
pctescolaspne = df_planilha['IN_ACESSIBILIDADE_INEXISTENTE'].value_counts(normalize = True).mul(100).apply(lambda x: "{:,.2f}".format(x))+'%'
print ("Quantidade de escolas que não possuem nenhuma forma de acessibilidade à pessoas com necessidades especiais: \n", qntescolaspne.to_string())
print ("Porcentagem de escolas que não possuem nenhuma forma de acessibilidade à pessoas com necessidades especiais: \n", pctescolaspne.to_string())
plt.title("Gráfico das escolas que não possuem nenhuma\n forma de acessibilidade à pessoas com necessidades especiais:")
plt.pie(qntescolaspne, autopct = lambda p : '{:.2f}%  ({:.0f})'.format(p, p * sum(qntescolaspne)/100))
plt.legend(["Possuem", "Não possuem"], loc = 'lower left')
plt.savefig('Imagens geradas dos gráficos/Gráfico de escolas que não possuem nenhuma forma de acessibilidade à pessoas com necessidades especiais.png')
plt.show()

print("Foi gerado um arquivo .png para cada gráfico analisado!")

# Para finaliazar o script, criei um novo DataFrame baseado nos cinco tópicos analisados, e pedi para que o módulo "openpyxl" criasse uma nova planilha baseada nesses dados.
nova_planilha = pd.DataFrame([[qntescolaslixo.to_string(), pctescolaslixo.to_string()], [qntescolasref.to_string(), pctescolasref.to_string()], 
[qntescolaslab.to_string(), pctescolaslab.to_string()], [qntescolasbanpne.to_string(), pctescolasbanpne.to_string()], [qntescolaspne.to_string(), 
pctescolaspne.to_string()]], index = ['Escolas que não realizam o tratamento do lixo', 'Escolas que possuem refeitório', 'Escolas que possuem um laboratório de informática',
 'Escolas que possuem banheiros acessíveis à pessoas com necessidades especiais','Escolas que não possuem nenhuma forma de acessibilidade à pessoas com necessidades especiais'], 
 columns = ['quantidade','porcentagem'])
nova_planilha.to_excel('Novas planilhas geradas/NovaPlanilha.xlsx', sheet_name = 'NovaPlanilha')
print("Uma nova planilha com os dados analisados foi gerada!")