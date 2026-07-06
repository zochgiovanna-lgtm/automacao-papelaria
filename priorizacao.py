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

# 3. Engenharia de Produção: Tempo Fixo (Setup) vs Tempo Unitário (por peça)
# Você pode ajustar esses números em minutos como achar melhor para a realidade da loja!
regras_tempos = {
    'Cartão de Visita':       {'fixo': 30, 'unitario': 0.5},    # 30 min fixos para criar a arte/ajustar o corte. Imprimir 10 ou 100 é o mesmo tempo.
    'Impressão':              {'fixo': 5,  'unitario': 0.1},  # 5 min para abrir o arquivo + 6 segundos (0.1 min) por folha impressa.
    'Convite':                {'fixo': 60, 'unitario': 2},    # 60 min de design + 2 min por unidade para dobrar e colar o laço.
    'Caneca':                 {'fixo': 0,  'unitario': 40},   # 40 min por caneca na prensa (produção linear).
    'Agenda':                 {'fixo': 0,  'unitario': 300},  # 5 horas por agenda (trabalho manual longo).
    'Camiseta':               {'fixo': 0,  'unitario': 20},   # 20 min por camiseta na prensa térmica.
    'Topo de Bolo':           {'fixo': 45, 'unitario': 15},   # 45 min criando o tema + 15 min montando cada unidade.
    'Balão Bubble Elaborado': {'fixo': 0,  'unitario': 130}   # 130 min por balão personalizado.
}

# 4. Tratamento de Segurança das Colunas Numéricas
tabela_pedidos['Dias_Para_Entrega'] = pd.to_numeric(tabela_pedidos['Dias_Para_Entrega'], errors='coerce').fillna(1)
tabela_pedidos['Dias_Para_Entrega'] = tabela_pedidos['Dias_Para_Entrega'].replace(0, 1)

tabela_pedidos['Quantidade'] = pd.to_numeric(tabela_pedidos['Quantidade'], errors='coerce').fillna(1)
tabela_pedidos['Quantidade'] = tabela_pedidos['Quantidade'].replace(0, 1)

# 5. Função matemática para calcular o tempo real com base no comportamento do produto
def calcular_tempo_real(linha):
    produto = linha['Produto']
    qtd = linha['Quantidade']
    
    if produto in regras_tempos:
        tempo_fixo = regras_tempos[produto]['fixo']
        tempo_unitario = regras_tempos[produto]['unitario']
        return tempo_fixo + (tempo_unitario * qtd)
    return 15 # Tempo padrão caso o produto não seja mapeado

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

# 7. Organizar por ordem de prioridade
tabela_pedidos['_nota_oculta'] = nota_interna
tabela_organizada = tabela_pedidos.sort_values(by='_nota_oculta', ascending=False)

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

print("Planilha atualizada com sucesso usando lógica de tempos industriais (Setup vs Unitário)!")
