#!/usr/bin/env python
# coding: utf-8

# Passo 1 - Importar Arquivos e Bibliotecas

#importar bibliotecas

import pandas as pd
import yagmail
import pathlib

#importar base de dados

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

# Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

#incluir nome da loja na planilha de vendas
vendas = vendas.merge(lojas, on='ID Loja')

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

dia_indicador = vendas['Data'].max()

# Passo 3 - Salvar a planilha na pasta de backup

#identificando se a pasta ja existe

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]    

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    
    #salvar dentro da pasta
    nome_arquivo = '{}_{}_{}.xlsx'. format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)


# Passo 4 - Calcular o indicador para 1 loja

#definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

for loja in dicionario_lojas:
    
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    #diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    #ticket medio
    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    
    #enviar e-mail
    usuario = yagmail.SMTP(user='gmail@gmail.com', password = '123@qwe')
    
    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    para = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    assunto = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    corpo_email = f'''
    <p>Bom dia, {nome}</p> 

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>

    <p>Segue anexo a planilha com todos os dados para mais detalhes.</p>
    <p>Qualquer dúvida estou a disposição.</p>
    <p>Att., Estagiário de dados</p>
    '''

    #Colocando anexos no e-mail
    anexo = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'

    #enviando e-mail

    usuario.send(to = 'gmail@gmail.com', subject = assunto, contents = corpo_email, attachments = str(anexo))
    print('E-mail da loja {} enviado'.format(loja))

# ### Passo 7 - Criar ranking para diretoria

faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)
display(faturamento_lojas)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'. format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
display(faturamento_lojas_dia)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'. format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

#enviar e-mail
usuario = yagmail.SMTP(user='gmail@gmail.com', password = '123@qwe')

para = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
assunto = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'

corpo_email = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Estagiário de dados
'''

#Colocando anexos no e-mail
anexo = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'


#enviando e-mail

usuario.send(to = 'gmail@gmail.com', subject = assunto, contents = corpo_email, attachments = str(anexo))
print('E-mail da Diretoria enviado')