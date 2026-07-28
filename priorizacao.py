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
# 3. GARANTIA DE COLUNAS EXISTENTES
# Isso impede o robô de quebrar caso falte alguma coluna na leitura
colunas_necessarias = ['Data_Pedido', 'Celular', 'Concluido', 'Observações']
for col in colunas_necessarias:
    if col not in tabela_pedidos.columns:
        tabela_pedidos[col] = ''
# ==========================================

# 4. Filtros de Limpeza Inteligente
# Remove linhas fantasmas (sem cliente)
tabela_pedidos = tabela_pedidos[tabela_pedidos['Cliente_ID'].astype(str).str.strip() != '']
# Tira da fila os pedidos concluídos (lê 'PRONTO', 'SIM', ou a caixinha de seleção marcada como 'TRUE')
tabela_pedidos = tabela_pedidos[~tabela_pedidos['Concluido'].astype(str).str.upper().isin(['TRUE', 'SIM', 'VERDADEIRO', 'PRONTO'])] 

# 5. Engenharia de Produção (com os produtos provisórios que ela inventou)
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
    'Envelope com Vale Presente':       {'fixo': 0, 'unitario': 30},
    'Foto Polaroid':                    {'fixo': 0, 'unitario': 5},
    'Foto Polaroid Imã de Geladeira':   {'fixo': 0, 'unitario': 5},
    'estampa dif':                      {'fixo': 0, 'unitario': 10},
    'Adesivi de vinil':                 {'fixo': 0, 'unitario': 5},
    'TAG AGRADECIMENTO':                {'fixo': 0, 'unitario': 5},
    'IMPRESSÃO DE CERTIFICADO':         {'fixo': 0, 'unitario': 5},
    'Tags Adesivo Personalizados':      {'fixo': 0, 'unitario': 10},
    'Foto polaroid de Geladeira':       {'fixo': 0, 'unitario': 5}
}

# 6. A Mágica do Calendário
tabela_pedidos['Data_Calc'] = pd.to_datetime(tabela_pedidos['Data_Entrega'], format='%d/%m/%Y', errors='coerce')
hoje = pd.Timestamp.today().normalize()
tabela_pedidos['Dias_Para_Entrega'] = (tabela_pedidos['Data_Calc'] - hoje).dt.days
tabela_pedidos['Dias_Para_Entrega'] = tabela_pedidos['Dias_Para_Entrega'].fillna(1)
tabela_pedidos['Dias_Para_Entrega'] = tabela_pedidos['Dias_Para_Entrega'].apply(lambda x: 1 if x <= 0 else x)

# 7. Proteção Master da Quantidade (Caça-Números)
# Se ela escrever "24 estampas", o robô arranca o texto e fica só com o "24"
tabela_pedidos['Quantidade_Limpa'] = tabela_pedidos['Quantidade'].astype(str).str.extract(r'(\d+)')[0]
tabela_pedidos['Quantidade_Limpa'] = pd.to_numeric(tabela_pedidos['Quantidade_Limpa'], errors='coerce').fillna(1)
tabela_pedidos['Quantidade_Limpa'] = tabela_pedidos['Quantidade_Limpa'].replace(0, 1)

# 8. Função de Tempo (Ignora se ela usar maiúsculas ou minúsculas)
def calcular_tempo_real(linha):
    produto = str(linha['Produto']).strip().lower()
    qtd = linha['Quantidade_Limpa']
    
    for chave_regra in regras_tempos:
        if chave_regra.lower() == produto:
            tempo_fixo = regras_tempos[chave_regra]['fixo']
            tempo_unitario = regras_tempos[chave_regra]['unitario']
            return tempo_fixo + (tempo_unitario * qtd)
    return 15

tabela_pedidos['Tempo_Total_Minutos'] = tabela_pedidos.apply(calcular_tempo_real, axis=1)

# 9. Cálculo da Urgência
nota_interna = tabela_pedidos['Tempo_Total_Minutos'] / tabela_pedidos['Dias_Para_Entrega']

def definir_status(nota):
    if nota >= 30:
        return '🚨 Urgente'
    elif nota >= 15:
        return '⚠️ Atenção'
    else:
        return '✅ Em dia'

tabela_pedidos['Status_Urgencia'] = nota_interna.apply(definir_status)

# 10. Organização por Prioridade para imprimir na planilha
tabela_pedidos['Quantidade'] = tabela_pedidos['Quantidade_Limpa'] # Passa o número limpo para a versão final
tabela_pedidos['_nota_oculta'] = nota_interna
tabela_organizada = tabela_pedidos.sort_values(by='_nota_oculta', ascending=False)

# Essa linha dita exatamente quais colunas vão aparecer na Fila_Prioridade e em qual ordem
colunas_finais = ['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Tempo_Total_Minutos', 'Status_Urgencia']
tabela_final_sheets = tabela_organizada[colunas_finais]

# 11. Enviar os dados de volta para a Fila
try:
    aba_destino = planilha.worksheet('Fila_Prioridade')
    aba_destino.clear() 
except:
    aba_destino = planilha.add_worksheet(title="Fila_Prioridade", rows="100", cols="20")

tabela_para_enviar = [tabela_final_sheets.columns.values.tolist()] + tabela_final_sheets.values.tolist()
aba_destino.update(tabela_para_enviar)

print("Sistema processado com sucesso! Filtros e proteções ativados.")
