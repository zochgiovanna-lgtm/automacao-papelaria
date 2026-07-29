import os
import json
import gspread
import pandas as pd
import numpy as np

# 1. Autenticação segura
credenciais_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
gc = gspread.service_account_from_dict(credenciais_json)

# 2. Abrir a planilha (Lendo a linha 2 como cabeçalho)
planilha = gc.open('Pedidos_Papelaria')
aba_origem = planilha.worksheet('Página1')
dados = aba_origem.get_all_records(head=2)
tabela_pedidos = pd.DataFrame(dados)

# 3. GARANTIA DE COLUNAS EXISTENTES
colunas_necessarias = ['Data_Pedido', 'Celular', 'Concluido', 'Observações', 'Quantidade']
for col in colunas_necessarias:
    if col not in tabela_pedidos.columns:
        tabela_pedidos[col] = ''

# 4. Limpeza de linhas fantasmas
tabela_pedidos = tabela_pedidos[tabela_pedidos['Cliente_ID'].astype(str).str.strip() != '']

# Separar Concluídos de Pendentes
condicao_concluido = tabela_pedidos['Concluido'].astype(str).str.upper().isin(['TRUE', 'SIM', 'VERDADEIRO', 'PRONTO'])
tabela_concluidos = tabela_pedidos[condicao_concluido].copy()
tabela_pendentes = tabela_pedidos[~condicao_concluido].copy()

# ==========================================
# PARTE A: ATUALIZAR ABA HISTÓRICO DE ENTREGAS
# ==========================================
colunas_historico = ['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Concluido']
for col in colunas_historico:
    if col not in tabela_concluidos.columns:
        tabela_concluidos[col] = ''
tabela_final_historico = tabela_concluidos[colunas_historico]

# Garantir que a aba Historico_Entregas existe, limpando APENAS os dados da linha 3 para baixo
try:
    aba_historico = planilha.worksheet('Historico_Entregas')
    try:
        aba_historico.batch_clear(['A3:H1000'])
    except:
        pass
except:
    aba_historico = planilha.add_worksheet(title="Historico_Entregas", rows="100", cols="20")

# Envia apenas os valores a partir da linha 3 (preservando totalmente a linha 2)
dados_historico = tabela_final_historico.values.tolist()
if dados_historico:
    aba_historico.update('A3', dados_historico)

# AUTOMATIZAÇÃO DA CAIXINHA DE ENTREGUES (J3 a L5 mescladas com emoji 🎉 e cor)
total_entregues = len(tabela_concluidos)
texto_contador = f"🎉 Pedidos Entregues: {total_entregues}"

try:
    # Insere o texto na célula J3
    aba_historico.update('J3', [[texto_contador]])
    
    # Mescla as colunas J, K, L e linhas 3 a 5 para formar o bloco/caixinha
    aba_historico.merge('J3:L5', merge_type='MERGE_ALL')
    
    # Formata a caixinha (Fundo rosa escuro/mauve, texto branco em negrito, centralizado)
    aba_historico.format('J3:L5', {
        'backgroundColor': {'red': 0.65, 'green': 0.38, 'blue': 0.48},
        'textFormat': {
            'bold': True, 
            'foregroundColor': {'red': 1.0, 'green': 1.0, 'blue': 1.0},
            'fontSize': 11
        },
        'horizontalAlignment': 'CENTER',
        'verticalAlignment': 'MIDDLE'
    })
except Exception as e:
    print(f"Erro ao atualizar caixinha de entregues: {e}")


# ==========================================
# PARTE B: PROCESSAR FILA DE PRIORIDADE (Apenas Pendentes)
# ==========================================
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
    'TAG AGRADECIMENTO':                {'fixo': 0, 'unitario': 5},
    'IMPRESSÃO DE CERTIFICADO':         {'fixo': 0, 'unitario': 5},
    'Tags Adesivo Personalizados':      {'fixo': 0, 'unitario': 10},
    'Foto polaroid de Geladeira':       {'fixo': 0, 'unitario': 5},
    'Estampa DTF':                      {'fixo': 0, 'unitario': 10},
    'Adesivo de Vinil':                 {'fixo': 0, 'unitario': 5}
}

if not tabela_pendentes.empty:
    tabela_pendentes['Data_Calc'] = pd.to_datetime(tabela_pendentes['Data_Entrega'], format='%d/%m/%Y', errors='coerce')
    hoje = pd.Timestamp.today().normalize()
    tabela_pendentes['Dias_Para_Entrega'] = (tabela_pendentes['Data_Calc'] - hoje).dt.days
    tabela_pendentes['Dias_Para_Entrega'] = tabela_pendentes['Dias_Para_Entrega'].fillna(1)
    tabela_pendentes['Dias_Para_Entrega'] = tabela_pendentes['Dias_Para_Entrega'].apply(lambda x: 1 if x <= 0 else x)

    tabela_pendentes['Quantidade_Limpa'] = tabela_pendentes['Quantidade'].astype(str).str.extract(r'(\d+)')[0]
    tabela_pendentes['Quantidade_Limpa'] = pd.to_numeric(tabela_pendentes['Quantidade_Limpa'], errors='coerce').fillna(1)
    tabela_pendentes['Quantidade_Limpa'] = tabela_pendentes['Quantidade_Limpa'].replace(0, 1)

    def calcular_tempo_real(linha):
        produto = str(linha['Produto']).strip().lower()
        qtd = linha['Quantidade_Limpa']
        for chave_regra in regras_tempos:
            if chave_regra.lower() == produto:
                tempo_fixo = regras_tempos[chave_regra]['fixo']
                tempo_unitario = regras_tempos[chave_regra]['unitario']
                return tempo_fixo + (tempo_unitario * qtd)
        return 15

    tabela_pendentes['Tempo_Total_Minutos'] = tabela_pendentes.apply(calcular_tempo_real, axis=1)

    nota_interna = tabela_pendentes['Tempo_Total_Minutos'] / tabela_pendentes['Dias_Para_Entrega']

    def definir_status(nota):
        if nota >= 30:
            return '🚨 Urgente'
        elif nota >= 15:
            return '⚠️ Atenção'
        else:
            return '✅ Em dia'

    tabela_pendentes['Status_Urgencia'] = nota_interna.apply(definir_status)
    tabela_pendentes['Quantidade'] = tabela_pendentes['Quantidade_Limpa']
    tabela_pendentes['_nota_oculta'] = nota_interna
    tabela_organizada = tabela_pendentes.sort_values(by='_nota_oculta', ascending=False)
else:
    tabela_organizada = pd.DataFrame(columns=['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Tempo_Total_Minutos', 'Status_Urgencia'])

colunas_finais = ['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Tempo_Total_Minutos', 'Status_Urgencia']
tabela_final_sheets = tabela_organizada[colunas_finais] if not tabela_organizada.empty else pd.DataFrame(columns=colunas_finais)

# Garantir que a aba Fila_Prioridade existe, limpando APENAS os dados da linha 3 para baixo
try:
    aba_destino = planilha.worksheet('Fila_Prioridade')
    try:
        aba_destino.batch_clear(['A3:I1000'])
    except:
        pass
except:
    aba_destino = planilha.add_worksheet(title="Fila_Prioridade", rows="100", cols="20")

# Envia apenas os valores a partir da linha 3 (preservando totalmente a linha 2)
dados_prioridade = tabela_final_sheets.values.tolist()
if dados_prioridade:
    aba_destino.update('A3', dados_prioridade)

print("Processamento concluído com sucesso! Caixinha formatada, mesclada com 🎉 e histórico atualizado.")
