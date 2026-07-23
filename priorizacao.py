import os
import json
import gspread
import pandas as pd
import numpy as np

# 1. Autenticação segura
credenciais_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
gc = gspread.service_account_from_dict(credenciais_json)

# 2. Abrir a planilha
planilha = gc.open('Pedidos_Papelaria')
aba_origem = planilha.worksheet('Página1')
dados = aba_origem.get_all_records()
tabela_pedidos = pd.DataFrame(dados)

# ==========================================
# NOVO: Filtro Anti-Linhas Fantasmas
# Remove qualquer linha onde a coluna 'Cliente_ID' esteja vazia
tabela_pedidos = tabela_pedidos[tabela_pedidos['Cliente_ID'].astype(str).str.strip() != '']
# ==========================================

# 3. Engenharia de Produção
regras_tempos = {
    'Cartão de Visita':                 {'fixo': 30, 'unitario': 0.5},
    'Impressão':                        {'fixo': 5,  'unitario': 0.1},
    'Caneca':                           {'fixo': 0,  'unitario': 40},
    'Agenda':                           {'fixo': 0,  'unitario': 300},
    'Camiseta':                         {'fixo': 0,  'unitario': 20},
    'Topo de Bolo':                     {'fixo': 45, 'unitario': 15},
    'Balão Bubble Elaborado':           {'fixo': 0,  'unitario': 130},
    'Papel Adesivo com Corte':          {'fixo': 0, 'unitario': 5},
    'Crachá Simples com Plastificação': {'fixo': 0, 'unitario': 15},
    'Convite Simples':                  {'fixo': 0, 'unitario': 20},
    'Convite Elaborado com Corte':      {'fixo': 0, 'unitario': 60},
    'Caixa Padrinho':                   {'fixo': 0, 'unitario': 60},
    'Card Simples':                     {'fixo': 0, 'unitario': 5},
    'Etiqueta Roupa':                   {'fixo': 0, 'unitario': 60},
    'Etiqueta Simples sem Laminação':   {'fixo': 0, 'unitario': 30},
    'Aplicação Nome Camiseta':          {'fixo': 0, 'unitario': 30},
    'Bloco de Pedidos':                 {'fixo': 0, 'unitario': 60},
    'Caderneta de Vacina Reforma':      {'fixo': 0, 'unitario': 240},
    'Card com Chocolate':               {'fixo': 0, 'unitario': 10},
    'Flay Simples':                     {'fixo': 0, 'unitario': 10},
    'Plastificação':                    {'fixo': 0, 'unitario': 10},
    'Reforma Agenda Escolar':           {'fixo': 0, 'unitario': 30},
    'Convite Casamento':                {'fixo': 4320, 'unitario': 0},
    'Corte Letras Color Pluss':         {'fixo': 0, 'unitario': 30},
    'Comanda':                          {'fixo': 4320, 'unitario': 0},
    'Apostila com Impressão':           {'fixo': 0, 'unitario': 240},
    'Envelope com Vale Presente':       {'fixo': 0, 'unitario': 30}
}

# 4. A Mágica do Calendário
tabela_pedidos['Data_Entrega'] = pd.to_datetime(tabela_pedidos['Data_Entrega'], format='%d/%m/%Y', errors='coerce')
hoje = pd.Timestamp.today().normalize()
tabela_pedidos['Dias_Para_Entrega'] = (tabela_pedidos['Data_Entrega'] - hoje).dt.days
tabela_pedidos['Dias_Para_Entrega'] = tabela_pedidos['Dias_Para_Entrega'].fillna(1)
tabela_pedidos['Dias_Para_Entrega'] = tabela_pedidos['Dias_Para_Entrega'].apply(lambda x: 1 if x <= 0 else x)

# Proteção da Quantidade
tabela_pedidos['Quantidade'] = pd.to_numeric(tabela_pedidos['Quantidade'], errors='coerce').fillna(1)
tabela_pedidos['Quantidade'] = tabela_pedidos['Quantidade'].replace(0, 1)

# 5. Função matemática para calcular o tempo real
def calcular_tempo_real(linha):
    produto = linha['Produto']
    qtd = linha['Quantidade']
    
    if produto in regras_tempos:
        tempo_fixo = regras_tempos[produto]['fixo']
        tempo_unitario = regras_tempos[produto]['unitario']
        return tempo_fixo + (tempo_unitario * qtd)
    return 15

tabela_pedidos['Tempo_Total_Minutos'] = tabela_pedidos.apply(calcular_tempo_real, axis=1)

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

# 7. Organização por Prioridade
tabela_pedidos['Data_Texto'] = tabela_pedidos['Data_Entrega'].dt.strftime('%d/%m/%Y').fillna('Sem Data')
tabela_pedidos['_nota_oculta'] = nota_interna
tabela_organizada = tabela_pedidos.sort_values(by='_nota_oculta', ascending=False)

colunas_finais = ['Cliente_ID', 'Produto', 'Quantidade', 'Data_Texto', 'Tempo_Total_Minutos', 'Status_Urgencia']
tabela_final_sheets = tabela_organizada[colunas_finais]
tabela_final_sheets = tabela_final_sheets.rename(columns={'Data_Texto': 'Data_Entrega'})

# 8. Enviar os dados de volta
try:
    aba_destino = planilha.worksheet('Fila_Prioridade')
    aba_destino.clear() 
except:
    aba_destino = planilha.add_worksheet(title="Fila_Prioridade", rows="100", cols="20")

tabela_para_enviar = [tabela_final_sheets.columns.values.tolist()] + tabela_final_sheets.values.tolist()
aba_destino.update(tabela_para_enviar)

print("Planilha atualizada sem as linhas fantasmas!")
