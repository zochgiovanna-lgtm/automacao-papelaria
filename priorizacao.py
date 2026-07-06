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

# 3. Dicionário com os tempos de produção BASE (por unidade ou lote)
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

# 4. Tratamento de Segurança das Colunas Numéricas
# Garante que os dias para entrega são números válidos
tabela_pedidos['Dias_Para_Entrega'] = pd.to_numeric(tabela_pedidos['Dias_Para_Entrega'], errors='coerce').fillna(1)
tabela_pedidos['Dias_Para_Entrega'] = tabela_pedidos['Dias_Para_Entrega'].replace(0, 1)

# Garante que a quantidade é um número válido (se estiver vazio, vira 1)
tabela_pedidos['Quantidade'] = pd.to_numeric(tabela_pedidos['Quantidade'], errors='coerce').fillna(1)
tabela_pedidos['Quantidade'] = tabela_pedidos['Quantidade'].replace(0, 1)

# 5. Cálculo do Tempo Total (Tempo Base x Quantidade)
tabela_pedidos['Tempo_Total_Minutos'] = tabela_pedidos['Produto'].map(tempos_produtos) * tabela_pedidos['Quantidade']

# 6. Cálculo da Urgência
nota_interna = tabela_pedidos['Tempo_Total_Minutos'] / tabela_pedidos['Dias_Para_Entrega']

def definir_status(nota):
    if nota >= 30:
        return '🚨 Urgente'
    elif nota >= 15:
        return '⚠️ Atenção'
    else:
        return '✅ Em dia'

tabela_pedidos['Status_Urgencia'] = nota_interna.apply(definir_status)

# 7. Organizar por ordem de prioridade
tabela_pedidos['_nota_oculta'] = nota_interna
tabela_organizada = tabela_pedidos.sort_values(by='_nota_oculta', ascending=False)

# Adicionamos a quantidade e o tempo total na folha final de prioridades
colunas_finais = ['Cliente_ID', 'Produto', 'Quantidade', 'Dias_Para_Entrega', 'Tempo_Total_Minutos', 'Status_Urgencia']
tabela_final_sheets = tabela_organizada[colunas_finais]

# 8. Enviar os dados de volta para a aba Fila_Prioridade
try:
    aba_destino = planilha.worksheet('Fila_Prioridade')
    aba_destino.clear() 
except:
    aba_destino = planilha.add_worksheet(title="Fila_Prioridade", rows="100", cols="20")

tabela_para_enviar = [tabela_final_sheets.columns.values.tolist()] + tabela_final_sheets.values.tolist()
aba_destino.update(tabela_para_enviar)

print("Planilha de prioridades atualizada com sucesso levando em conta as quantidades!")

tabela_para_enviar = [tabela_final_sheets.columns.values.tolist()] + tabela_final_sheets.values.tolist()
aba_destino.update(tabela_para_enviar)

print("Planilha atualizada com sucesso pelo servidor em nuvem!")
