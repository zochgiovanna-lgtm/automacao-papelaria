import os
import json
import gspread
import pandas as pd
import numpy as np

# 1. Autenticação segura usando a chave que guardamos no GitHub Secrets
credenciais_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
gc = gspread.service_account_from_dict(credenciais_json)

# 2. Abrir a planilha e puxar os dados da Página1
planilha = gc.open('Pedidos_Papelaria')
aba_origem = planilha.worksheet('Página1')
dados = aba_origem.get_all_records()
tabela_pedidos = pd.DataFrame(dados)

# 3. Dicionário com os tempos de produção (em minutos) - COM O BALÃO INCLUSO!
tempos_produtos = {
    'Cartão de Visita': 30,
    'Caneca': 40,
    'Convite': 60,
    'Topo de Bolo': 90,
    'Agenda': 300,
    'Camiseta': 20,       
    'Impressão': 5,
    'Balão Bubble Elaborado': 130
}

# 4. Processamento dos dados e cálculo de urgência
tabela_pedidos['Tempo_Producao_Minutos'] = tabela_pedidos['Produto'].map(tempos_produtos)

tabela_pedidos['Dias_Para_Entrega'] = pd.to_numeric(tabela_pedidos['Dias_Para_Entrega'], errors='coerce').fillna(1)
tabela_pedidos['Dias_Para_Entrega'] = tabela_pedidos['Dias_Para_Entrega'].replace(0, 1)

nota_interna = tabela_pedidos['Tempo_Producao_Minutos'] / tabela_pedidos['Dias_Para_Entrega']

def definir_status(nota):
    if nota >= 30:
        return '🚨 URGENTE'
    elif nota >= 15:
        return '⚠️ Prioridade'
    else:
        return '✅ Em dia'

tabela_pedidos['Status_Urgencia'] = nota_interna.apply(definir_status)

tabela_pedidos['_nota_oculta'] = nota_interna
tabela_organizada = tabela_pedidos.sort_values(by='_nota_oculta', ascending=False)

colunas_finais = ['Cliente_ID', 'Produto', 'Dias_Para_Entrega', 'Tempo_Producao_Minutos', 'Status_Urgencia']
tabela_final_sheets = tabela_organizada[colunas_finais]

# 5. Enviar os dados de volta para a aba Fila_Prioridade
try:
    aba_destino = planilha.worksheet('Fila_Prioridade')
    aba_destino.clear() 
except:
    aba_destino = planilha.add_worksheet(title="Fila_Prioridade", rows="100", cols="20")

tabela_para_enviar = [tabela_final_sheets.columns.values.tolist()] + tabela_final_sheets.values.tolist()
aba_destino.update(tabela_para_enviar)

print("Planilha atualizada com sucesso pelo servidor em nuvem!")
