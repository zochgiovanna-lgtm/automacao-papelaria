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
colunas_necessarias = ['Data_Pedido', 'Celular', 'Concluido', 'Observações', 'Quantidade', 'Pago']
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
colunas_historico = ['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Concluido', 'Pago']
for col in colunas_historico:
    if col not in tabela_concluidos.columns:
        tabela_concluidos[col] = ''
tabela_final_historico = tabela_concluidos[colunas_historico]

try:
    aba_historico = planilha.worksheet('Historico_Entregas')
    try:
        aba_historico.batch_clear(['A3:I1000'])
    except Exception:
        pass
except gspread.exceptions.WorksheetNotFound:
    aba_historico = planilha.add_worksheet(title="Historico_Entregas", rows="100", cols="20")

dados_historico = tabela_final_historico.values.tolist()
if dados_historico:
    aba_historico.update('A3', dados_historico)

total_entregues = len(tabela_concluidos)
texto_contador = f"🎉 Pedidos Entregues: {total_entregues}"

try:
    aba_historico.update('J1', [[texto_contador]])
except Exception as e:
    print(f"Erro ao atualizar contador: {e}")


# ==========================================
# PARTE B: PROCESSAR FILA DE PRIORIDADE (Apenas Pendentes)
# ==========================================
regras_tempos = {
    'CARTÃO DE VISITA':                 {'fixo': 30, 'unitario': 0.5},
    'IMPRESSÃO':                        {'fixo': 5,  'unitario': 0.1},
    'CANECA':                           {'fixo': 0,  'unitario': 40},
    'AGENDA':                           {'fixo': 0,  'unitario': 300},
    'CAMISETA':                         {'fixo': 0,  'unitario': 20},
    'TOPO DE BOLO':                     {'fixo': 45, 'unitario': 15},
    'BALÃO BUBBLE ELABORADO':           {'fixo': 0,  'unitario': 130},
    'PAPEL ADESIVO COM CORTE':          {'fixo': 0, 'unitario': 5},
    'CRACHÁ SIMPLES COM PLASTIFICAÇÃO': {'fixo': 0, 'unitario': 15},
    'CONVITE SIMPLES':                  {'fixo': 0, 'unitario': 20},
    'CONVITE ELABORADO COM CORTE':      {'fixo': 0, 'unitario': 60},
    'CAIXA PADRINHO':                   {'fixo': 0, 'unitario': 60},
    'CARD SIMPLES':                     {'fixo': 0, 'unitario': 5},
    'ETIQUETA ROUPA':                   {'fixo': 0, 'unitario': 60},
    'ETIQUETA SIMPLES SEM LAMINAÇÃO':   {'fixo': 0, 'unitario': 30},
    'APLICAÇÃO NOME CAMISETA':          {'fixo': 0, 'unitario': 30},
    'BLOCO DE PEDIDOS':                 {'fixo': 0, 'unitario': 60},
    'CADERNETA DE VACINA REFORMA':      {'fixo': 0, 'unitario': 240},
    'CARD COM CHOCOLATE':               {'fixo': 0, 'unitario': 10},
    'FLAY SIMPLES':                     {'fixo': 0, 'unitario': 10},
    'PLASTIFICAÇÃO':                    {'fixo': 0, 'unitario': 10},
    'REFORMA AGENDA ESCOLAR':           {'fixo': 0, 'unitario': 30},
    'CONVITE CASAMENTO':                {'fixo': 4320, 'unitario': 0},
    'CORTE LETRAS COLOR PLUSS':         {'fixo': 0, 'unitario': 30},
    'COMANDA':                          {'fixo': 4320, 'unitario': 0},
    'APOSTILA COM IMPRESSÃO':           {'fixo': 0, 'unitario': 240},
    'ENVELOPE COM VALE PRESENTE':       {'fixo': 0, 'unitario': 30},
    'VALE PRESENTE':                    {'fixo': 0, 'unitario': 30},
    'FOTO POLAROID':                    {'fixo': 0, 'unitario': 5},
    'FOTO POLAROID IMÃ DE GELADEIRA':   {'fixo': 0, 'unitario': 5},
    'TAG AGRADECIMENTO':                {'fixo': 0, 'unitario': 5},
    'IMPRESSÃO DE CERTIFICADO':         {'fixo': 0, 'unitario': 5},
    'TAGS ADESIVO PERSONALIZADOS':      {'fixo': 0, 'unitario': 10},
    'FOTO POLAROID DE GELADEIRA':       {'fixo': 0, 'unitario': 5},
    'ESTAMPA DTF':                      {'fixo': 0, 'unitario': 10},
    'ADESIVO DE VINIL':                 {'fixo': 0, 'unitario': 5},
    'BLOQUINHO COLOR PLUS':             {'fixo': 0, 'unitario': 30}
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
    tabela_organizada = pd.DataFrame(columns=['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Tempo_Total_Minutos', 'Status_Urgencia', 'Pago'])

colunas_finais = ['Data_Pedido', 'Cliente_ID', 'Celular', 'Produto', 'Quantidade', 'Observações', 'Data_Entrega', 'Tempo_Total_Minutos', 'Status_Urgencia', 'Pago']
tabela_final_sheets = tabela_organizada[colunas_finais] if not tabela_organizada.empty else pd.DataFrame(columns=colunas_finais)

try:
    aba_destino = planilha.worksheet('Fila_Prioridade')
    try:
        aba_destino.batch_clear(['A3:J1000'])
    except Exception:
        pass
except gspread.exceptions.WorksheetNotFound:
    aba_destino = planilha.add_worksheet(title="Fila_Prioridade", rows="100", cols="20")

dados_prioridade = tabela_final_sheets.values.tolist()
if dados_prioridade:
    aba_destino.update('A3', dados_prioridade)


# ==========================================
# PARTE C: DISTRIBUIR PEDIDOS PELAS ABAS DE DIA DA SEMANA
# ==========================================
CAPACIDADE_DIARIA_MINUTOS = 480  # 8 horas — ajuste aqui se necessário

NOMES_DIAS = {
    0: 'SEGUNDA FEIRA',
    1: 'TERÇA FEIRA',
    2: 'QUARTA FEIRA',
    3: 'QUINTA FEIRA',
    4: 'SEXTA FEIRA',
    5: 'SÁBADO',
}

if not tabela_organizada.empty:
    dia_semana_hoje = pd.Timestamp.today().weekday()
    inicio = dia_semana_hoje if dia_semana_hoje < 6 else 0
    ordem_dias = [(inicio + i) % 6 for i in range(6)]
    nomes_em_ordem = [NOMES_DIAS[d] for d in ordem_dias]

    capacidade_restante = {nome: CAPACIDADE_DIARIA_MINUTOS for nome in nomes_em_ordem}
    pedidos_por_dia = {nome: [] for nome in nomes_em_ordem}

    for _, linha in tabela_organizada.iterrows():
        tempo = linha['Tempo_Total_Minutos']
        dia_escolhido = next(
            (nome for nome in nomes_em_ordem if capacidade_restante[nome] >= tempo),
            None
        )
        if dia_escolhido is None:
            dia_escolhido = max(capacidade_restante, key=capacidade_restante.get)

        capacidade_restante[dia_escolhido] -= tempo
        pedidos_por_dia[dia_escolhido].append(linha)

    for nome_aba in nomes_em_ordem:
        try:
            aba_dia = planilha.worksheet(nome_aba)
            try:
                aba_dia.batch_clear(['A3:J1000'])
            except Exception:
                pass
        except gspread.exceptions.WorksheetNotFound:
            aba_dia = planilha.add_worksheet(title=nome_aba, rows="100", cols="20")

        linhas_do_dia = pedidos_por_dia[nome_aba]
        if linhas_do_dia:
            tabela_dia = pd.DataFrame(linhas_do_dia)[colunas_finais]
            aba_dia.update('A3', tabela_dia.values.tolist())

    print("Distribuição por dia da semana concluída!")
else:
    print("Nenhum pedido pendente para distribuir pelos dias da semana.")

print("Processamento concluído com sucesso!")
