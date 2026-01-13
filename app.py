import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from datetime import datetime, date
import numpy as np

# Configuração da página
st.set_page_config(
    page_title="📊 Relatório de Cursos",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A5F;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #4A6FA5;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .progress-green { background-color: #28a745; }
    .progress-yellow { background-color: #ffc107; }
    .progress-red { background-color: #dc3545; }
    .status-concluido { color: #28a745; font-weight: bold; }
    .status-andamento { color: #ffc107; font-weight: bold; }
    .status-pendente { color: #dc3545; font-weight: bold; }
    .storytelling-box {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        border-left: 5px solid #1E3A5F;
        margin-bottom: 2rem;
    }
    .highlight-box {
        background: #fff3cd;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #ffc107;
    }
    .success-box {
        background: #d4edda;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #28a745;
    }
    .danger-box {
        background: #f8d7da;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #dc3545;
    }
</style>
""", unsafe_allow_html=True)


def load_data(uploaded_file):
    """Carrega os dados do arquivo Excel"""
    xl = pd.ExcelFile(uploaded_file)
    
    # Identifica as abas (pode ser Plano/Real ou Plano/Realizado)
    sheet_names = xl.sheet_names
    
    df_plano = pd.read_excel(xl, 'Plano')
    
    # Tenta encontrar a aba de realizados
    real_sheet = 'Real' if 'Real' in sheet_names else 'Realizado'
    df_real = pd.read_excel(xl, real_sheet)
    
    return df_plano, df_real


def process_data(df_plano, df_real):
    """Processa e agrega os dados"""
    
    # Normaliza a coluna de finalização
    df_real['Status'] = df_real['Finalizou o curso?'].apply(
        lambda x: 'Concluído' if str(x).lower() in ['sim', 'yes', 's'] else 
                  ('Em Andamento' if str(x).lower() in ['em andamento', 'andamento', 'in progress'] else 'Pendente')
    )
    
    # Verifica se tem data de início para classificar como "Em Andamento"
    if 'Data de início' in df_real.columns:
        df_real['Status'] = df_real.apply(
            lambda row: 'Em Andamento' if (row['Status'] == 'Pendente' and 
                                           pd.notna(row['Data de início']) and 
                                           str(row['Data de início']) != '-') else row['Status'],
            axis=1
        )
    
    # Calcula horas realizadas por colaborador
    df_real['Horas_Realizadas'] = df_real.apply(
        lambda row: row['Carga Horária'] if row['Status'] == 'Concluído' else 0,
        axis=1
    )
    
    # Agrupa por colaborador
    horas_realizadas = df_real.groupby(['Id colaborador(a)', 'Colaborador(a)'])['Horas_Realizadas'].sum().reset_index()
    
    # Merge com plano
    df_merged = pd.merge(
        df_plano,
        horas_realizadas,
        on=['Id colaborador(a)', 'Colaborador(a)'],
        how='left'
    )
    df_merged['Horas_Realizadas'] = df_merged['Horas_Realizadas'].fillna(0)
    df_merged['Percentual'] = (df_merged['Horas_Realizadas'] / df_merged['horas totais'] * 100).round(1)
    df_merged['Horas_Pendentes'] = df_merged['horas totais'] - df_merged['Horas_Realizadas']
    
    return df_merged, df_real


def create_bar_chart(df_merged):
    """Cria gráfico de barras horizontais comparando planejado vs realizado"""
    df_sorted = df_merged.sort_values('Percentual', ascending=True)
    
    fig = go.Figure()
    
    # Barras de horas planejadas (fundo)
    fig.add_trace(go.Bar(
        y=df_sorted['Colaborador(a)'],
        x=df_sorted['horas totais'],
        name='Planejado',
        orientation='h',
        marker_color='#E8E8E8',
        text=df_sorted['horas totais'].astype(int).astype(str) + 'h',
        textposition='outside',
        textfont=dict(size=10)
    ))
    
    # Barras de horas realizadas
    fig.add_trace(go.Bar(
        y=df_sorted['Colaborador(a)'],
        x=df_sorted['Horas_Realizadas'],
        name='Realizado',
        orientation='h',
        marker_color=df_sorted['Percentual'].apply(
            lambda x: '#28a745' if x >= 70 else ('#ffc107' if x >= 30 else '#dc3545')
        ),
        text=df_sorted['Horas_Realizadas'].astype(int).astype(str) + 'h (' + df_sorted['Percentual'].astype(str) + '%)',
        textposition='inside',
        textfont=dict(size=10, color='white')
    ))
    
    fig.update_layout(
        title='📊 Horas Planejadas vs Realizadas por Colaborador',
        barmode='overlay',
        height=400,
        xaxis_title='Horas',
        yaxis_title='',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        margin=dict(l=10, r=10, t=60, b=40)
    )
    
    return fig


def create_pie_chart(df_merged):
    """Cria gráfico de pizza com percentual geral de conclusão"""
    total_planejado = df_merged['horas totais'].sum()
    total_realizado = df_merged['Horas_Realizadas'].sum()
    total_pendente = total_planejado - total_realizado
    percentual_geral = (total_realizado / total_planejado * 100)
    
    fig = go.Figure(data=[go.Pie(
        labels=['Concluído', 'Pendente'],
        values=[total_realizado, total_pendente],
        hole=0.6,
        marker_colors=['#28a745', '#E8E8E8'],
        textinfo='percent+label',
        textfont=dict(size=12)
    )])
    
    fig.add_annotation(
        text=f'<b>{percentual_geral:.1f}%</b><br>Concluído',
        x=0.5, y=0.5,
        font=dict(size=18, color='#1E3A5F'),
        showarrow=False
    )
    
    fig.update_layout(
        title='🎯 Progresso Geral do Time',
        height=350,
        margin=dict(l=10, r=10, t=60, b=10),
        showlegend=True,
        legend=dict(orientation='h', yanchor='bottom', y=-0.1, xanchor='center', x=0.5)
    )
    
    return fig, percentual_geral, total_realizado, total_planejado


def calcular_dias_uteis_2026(data_inicio, data_fim):
    """Calcula dias úteis entre duas datas, descontando fins de semana e feriados nacionais de 2026"""
    
    # Feriados nacionais de 2026 (Brasil)
    feriados_2026 = [
        date(2026, 1, 1),   # Confraternização Universal
        date(2026, 2, 16),  # Carnaval (segunda)
        date(2026, 2, 17),  # Carnaval (terça)
        date(2026, 2, 18),  # Quarta de cinzas (ponto facultativo, mas muitas empresas emendam)
        date(2026, 4, 3),   # Sexta-feira Santa
        date(2026, 4, 21),  # Tiradentes
        date(2026, 5, 1),   # Dia do Trabalho
        date(2026, 6, 4),   # Corpus Christi
        date(2026, 9, 7),   # Independência do Brasil
        date(2026, 10, 12), # Nossa Senhora Aparecida
        date(2026, 11, 2),  # Finados
        date(2026, 11, 15), # Proclamação da República
        date(2026, 12, 25), # Natal
    ]
    
    dias_uteis = 0
    data_atual = data_inicio
    
    while data_atual <= data_fim:
        # Verifica se não é fim de semana (0=segunda, 6=domingo)
        if data_atual.weekday() < 5:  # Segunda a sexta
            # Verifica se não é feriado
            if data_atual not in feriados_2026:
                dias_uteis += 1
        data_atual += pd.Timedelta(days=1)
    
    return dias_uteis


def create_pace_chart(df_merged):
    """Cria gráfico de ritmo necessário para cada colaborador cumprir o prazo"""
    
    # Configurações
    data_atual = date.today()
    data_limite = date(2026, 12, 20)
    
    # Calcula dias totais e dias úteis
    dias_totais = (data_limite - data_atual).days
    dias_uteis_total = calcular_dias_uteis_2026(data_atual, data_limite)
    
    # Considera apenas 70% dos dias úteis (margem para imprevistos, reuniões, etc.)
    dias_uteis = int(dias_uteis_total * 0.70)
    
    # Calcula ritmo necessário para cada colaborador
    df_pace = df_merged.copy()
    df_pace['Horas_Restantes'] = df_pace['horas totais'] - df_pace['Horas_Realizadas']
    df_pace['Ritmo_Necessario'] = (df_pace['Horas_Restantes'] / dias_uteis).round(2)
    
    # Calcula ritmo ideal (horas totais / dias efetivos - o que deveria fazer desde o início)
    df_pace['Ritmo_Ideal'] = (df_pace['horas totais'] / dias_uteis).round(2)
    
    # Classifica o status por valores fixos
    def classify_status(ritmo):
        if ritmo <= 0:
            return '✅ Concluído'
        elif ritmo <= 1:
            return '🔵 Tranquilo'
        elif ritmo <= 1.5:
            return '🟢 Bom Ritmo'
        elif ritmo <= 2:
            return '🟡 Atenção'
        elif ritmo <= 3:
            return '🟠 Crítico'
        else:
            return '🔴 Plano de Ação'
    
    df_pace['Status_Ritmo'] = df_pace['Ritmo_Necessario'].apply(classify_status)
    
    # Ordena pelo ritmo necessário (mais crítico primeiro)
    df_pace = df_pace.sort_values('Ritmo_Necessario', ascending=True)
    
    # Define cores baseadas no ritmo (valores fixos)
    colors = df_pace['Ritmo_Necessario'].apply(
        lambda x: '#28a745' if x <= 0 else ('#3498db' if x <= 1 else ('#2ecc71' if x <= 1.5 else ('#f1c40f' if x <= 2 else ('#e67e22' if x <= 3 else '#e74c3c'))))
    ).tolist()
    
    fig = go.Figure()
    
    # Barras do ritmo necessário atual
    fig.add_trace(go.Bar(
        y=df_pace['Colaborador(a)'],
        x=df_pace['Ritmo_Necessario'],
        orientation='h',
        name='Ritmo Necessário',
        marker_color=colors,
        text=df_pace.apply(lambda row: f"{row['Ritmo_Necessario']:.1f}h/dia" if row['Ritmo_Necessario'] > 0 else "✅", axis=1),
        textposition='outside',
        textfont=dict(size=11, color='#333'),
        hovertemplate='<b>%{y}</b><br>Ritmo necessário: %{x:.2f}h/dia<br>Horas restantes: %{customdata[0]:.0f}h<br>Ritmo ideal: %{customdata[1]:.2f}h/dia<extra></extra>',
        customdata=df_pace[['Horas_Restantes', 'Ritmo_Ideal']].values
    ))
    
    # Linhas de referência fixas
    fig.add_vline(x=1.5, line_dash="dash", line_color="#2ecc71", line_width=2, 
                  annotation_text="1.5h/dia (ideal)", annotation_position="top")
    fig.add_vline(x=2, line_dash="dash", line_color="#e74c3c", line_width=2,
                  annotation_text="2h/dia (máx)", annotation_position="top")
    
    fig.update_layout(
        title=f'⏱️ Ritmo Necessário para Concluir até 20/12/2026<br><sub>Restam {dias_totais} dias corridos | {dias_uteis_total} dias úteis | <b>{dias_uteis} dias efetivos (70%)</b></sub>',
        height=400,
        xaxis_title='Horas por dia efetivo necessárias',
        yaxis_title='',
        xaxis=dict(range=[0, max(df_pace['Ritmo_Necessario'].max() * 1.3, df_pace['Ritmo_Ideal'].max() * 1.3, 3)]),
        margin=dict(l=10, r=80, t=80, b=40),
        showlegend=False
    )
    
    return fig, df_pace, dias_totais, dias_uteis, dias_uteis_total


def create_gauge_chart(percentual, nome):
    """Cria gráfico de gauge para progresso individual"""
    color = '#28a745' if percentual >= 70 else ('#ffc107' if percentual >= 30 else '#dc3545')
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=percentual,
        number={'suffix': '%', 'font': {'size': 24}},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1},
            'bar': {'color': color},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "gray",
            'steps': [
                {'range': [0, 30], 'color': '#ffebee'},
                {'range': [30, 70], 'color': '#fff8e1'},
                {'range': [70, 100], 'color': '#e8f5e9'}
            ],
            'threshold': {
                'line': {'color': "black", 'width': 2},
                'thickness': 0.75,
                'value': percentual
            }
        }
    ))
    
    fig.update_layout(
        height=200,
        margin=dict(l=20, r=20, t=30, b=20)
    )
    
    return fig


def create_status_table(df_real, colaborador):
    """Cria tabela de status dos cursos por colaborador"""
    df_colab = df_real[df_real['Colaborador(a)'] == colaborador].copy()
    
    # Ordena por status
    status_order = {'Concluído': 0, 'Em Andamento': 1, 'Pendente': 2}
    df_colab['Status_Order'] = df_colab['Status'].map(status_order)
    df_colab = df_colab.sort_values('Status_Order')
    
    return df_colab[['Curso', 'Carga Horária', 'Status']]


def get_status_icon(status):
    """Retorna ícone baseado no status"""
    if status == 'Concluído':
        return '✅'
    elif status == 'Em Andamento':
        return '🔄'
    else:
        return '❌'


def get_status_color(status):
    """Retorna cor baseada no status"""
    if status == 'Concluído':
        return '#28a745'
    elif status == 'Em Andamento':
        return '#ffc107'
    else:
        return '#dc3545'


def generate_pdf_content(df_merged, df_real, percentual_geral, total_realizado, total_planejado):
    """Gera conteúdo HTML para PDF com gráficos"""
    
    # Encontra melhores e piores desempenhos
    melhor = df_merged.loc[df_merged['Percentual'].idxmax()]
    pior = df_merged.loc[df_merged['Percentual'].idxmin()]
    
    # Calcula dados de ritmo
    data_atual = date.today()
    data_limite = date(2026, 12, 20)
    dias_totais = (data_limite - data_atual).days
    dias_uteis_total = calcular_dias_uteis_2026(data_atual, data_limite)
    dias_uteis = int(dias_uteis_total * 0.70)  # 70% dos dias úteis (margem para imprevistos)
    
    # Prepara dados de ritmo
    df_pace = df_merged.copy()
    df_pace['Horas_Restantes'] = df_pace['horas totais'] - df_pace['Horas_Realizadas']
    df_pace['Ritmo_Necessario'] = (df_pace['Horas_Restantes'] / dias_uteis).round(2)
    df_pace = df_pace.sort_values('Ritmo_Necessario', ascending=False)
    
    # Contagem de status
    cursos_concluidos = len(df_real[df_real['Status'] == 'Concluído'])
    cursos_andamento = len(df_real[df_real['Status'] == 'Em Andamento'])
    cursos_pendentes = len(df_real[df_real['Status'] == 'Pendente'])
    total_cursos = len(df_real)
    
    # Críticos (> 2h/dia)
    criticos = len(df_pace[df_pace['Ritmo_Necessario'] > 2])
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{ size: A4; margin: 1.2cm; }}
            body {{ font-family: Arial, sans-serif; font-size: 10px; line-height: 1.3; color: #333; }}
            .header {{ text-align: center; margin-bottom: 15px; border-bottom: 3px solid #1E3A5F; padding-bottom: 8px; }}
            .header h1 {{ color: #1E3A5F; margin: 0; font-size: 20px; }}
            .header p {{ color: #666; margin: 3px 0 0 0; font-size: 11px; }}
            .section {{ margin-bottom: 12px; }}
            .section-title {{ background: #1E3A5F; color: white; padding: 6px 10px; font-size: 12px; font-weight: bold; margin-bottom: 8px; border-radius: 4px; }}
            .storytelling {{ background: #f5f7fa; padding: 10px; border-left: 4px solid #1E3A5F; margin-bottom: 12px; font-size: 10px; }}
            .metrics-row {{ display: flex; justify-content: space-between; margin-bottom: 12px; gap: 8px; }}
            .metric-box {{ text-align: center; padding: 8px; border-radius: 6px; flex: 1; }}
            .metric-box.blue {{ background: #e3f2fd; border: 2px solid #1976d2; }}
            .metric-box.green {{ background: #e8f5e9; border: 2px solid #28a745; }}
            .metric-box.orange {{ background: #fff3e0; border: 2px solid #ff9800; }}
            .metric-box.red {{ background: #ffebee; border: 2px solid #dc3545; }}
            .metric-box.purple {{ background: #f3e5f5; border: 2px solid #9c27b0; }}
            .metric-value {{ font-size: 18px; font-weight: bold; }}
            .metric-label {{ font-size: 9px; color: #666; }}
            table {{ width: 100%; border-collapse: collapse; font-size: 9px; margin-bottom: 8px; }}
            th {{ background: #1E3A5F; color: white; padding: 5px; text-align: left; }}
            td {{ padding: 4px; border-bottom: 1px solid #ddd; }}
            .status-green {{ color: #28a745; font-weight: bold; }}
            .status-yellow {{ color: #ff9800; font-weight: bold; }}
            .status-red {{ color: #dc3545; font-weight: bold; }}
            .progress-bar {{ width: 100%; height: 12px; background: #e0e0e0; border-radius: 6px; overflow: hidden; }}
            .progress-fill {{ height: 100%; border-radius: 6px; }}
            .highlight {{ display: flex; gap: 10px; margin-bottom: 12px; }}
            .highlight-box {{ flex: 1; padding: 8px; border-radius: 6px; font-size: 10px; }}
            .highlight-box.success {{ background: #d4edda; border-left: 4px solid #28a745; }}
            .highlight-box.danger {{ background: #f8d7da; border-left: 4px solid #dc3545; }}
            .highlight-box.warning {{ background: #fff3cd; border-left: 4px solid #ffc107; }}
            .page-break {{ page-break-before: always; }}
            .two-col {{ display: flex; gap: 15px; }}
            .two-col > div {{ flex: 1; }}
            .chart-container {{ margin-bottom: 12px; }}
            .bar-chart {{ width: 100%; }}
            .bar-row {{ display: flex; align-items: center; margin-bottom: 6px; }}
            .bar-label {{ width: 140px; font-size: 9px; font-weight: 500; }}
            .bar-container {{ flex: 1; height: 18px; background: #e0e0e0; border-radius: 4px; position: relative; overflow: hidden; }}
            .bar-fill {{ height: 100%; border-radius: 4px; display: flex; align-items: center; justify-content: flex-end; padding-right: 5px; }}
            .bar-text {{ font-size: 8px; color: white; font-weight: bold; }}
            .bar-value {{ width: 70px; text-align: right; font-size: 9px; font-weight: bold; margin-left: 8px; }}
            .pie-container {{ display: flex; justify-content: center; align-items: center; gap: 20px; }}
            .pie-chart {{ width: 120px; height: 120px; border-radius: 50%; position: relative; }}
            .pie-legend {{ font-size: 10px; }}
            .pie-legend-item {{ display: flex; align-items: center; gap: 5px; margin-bottom: 4px; }}
            .legend-color {{ width: 12px; height: 12px; border-radius: 3px; }}
            .ritmo-bar {{ display: flex; align-items: center; margin-bottom: 4px; }}
            .ritmo-name {{ width: 130px; font-size: 9px; }}
            .ritmo-container {{ flex: 1; height: 16px; background: #f0f0f0; border-radius: 4px; position: relative; }}
            .ritmo-fill {{ height: 100%; border-radius: 4px; }}
            .ritmo-value {{ width: 60px; text-align: right; font-size: 9px; font-weight: bold; }}
            .ritmo-line {{ position: absolute; top: 0; bottom: 0; width: 2px; z-index: 10; }}
            .colaborador-section {{ margin-bottom: 12px; padding: 8px; border: 1px solid #ddd; border-radius: 6px; page-break-inside: avoid; }}
            .colaborador-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }}
            .colaborador-name {{ font-size: 11px; font-weight: bold; color: #1E3A5F; }}
            .info-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-bottom: 10px; }}
            .info-item {{ text-align: center; padding: 5px; background: #f8f9fa; border-radius: 4px; }}
            .info-value {{ font-size: 14px; font-weight: bold; }}
            .info-label {{ font-size: 8px; color: #666; }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>📊 Relatório de Acompanhamento de Cursos</h1>
            <p>Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')} | Prazo: 20/12/2026</p>
        </div>
        
        <div class="section">
            <div class="storytelling">
                <strong>🎯 Contexto:</strong> Plano de desenvolvimento focado em <b>liderança, estatística, dados e ferramentas digitais</b>.
                <strong>📈 Status:</strong> <b>{int(total_realizado)}h</b> de <b>{int(total_planejado)}h</b> concluídas (<b>{percentual_geral:.1f}%</b>).
                <strong>⏱️ Prazo:</strong> {dias_totais} dias corridos | <b>{dias_uteis} dias efetivos</b> (70% dos dias úteis, considerando imprevistos).
            </div>
        </div>

        <div class="section">
            <div class="section-title">📊 RESUMO EXECUTIVO</div>
            
            <div class="metrics-row">
                <div class="metric-box blue">
                    <div class="metric-value">{len(df_merged)}</div>
                    <div class="metric-label">Colaboradores</div>
                </div>
                <div class="metric-box green">
                    <div class="metric-value">{int(total_realizado)}h</div>
                    <div class="metric-label">Concluídas</div>
                </div>
                <div class="metric-box orange">
                    <div class="metric-value">{int(total_planejado - total_realizado)}h</div>
                    <div class="metric-label">Pendentes</div>
                </div>
                <div class="metric-box {'green' if percentual_geral >= 50 else 'red'}">
                    <div class="metric-value">{percentual_geral:.1f}%</div>
                    <div class="metric-label">Progresso</div>
                </div>
                <div class="metric-box purple">
                    <div class="metric-value">{dias_uteis}</div>
                    <div class="metric-label">Dias Úteis</div>
                </div>
                <div class="metric-box {'green' if criticos == 0 else 'red'}">
                    <div class="metric-value">{criticos}</div>
                    <div class="metric-label">Críticos</div>
                </div>
            </div>

            <div class="highlight">
                <div class="highlight-box success">
                    <strong>🏆 Melhor:</strong> {melhor['Colaborador(a)']}<br>
                    <span style="font-size: 14px; color: #28a745;"><b>{melhor['Percentual']:.1f}%</b></span> ({int(melhor['Horas_Realizadas'])}h/{int(melhor['horas totais'])}h)
                </div>
                <div class="highlight-box danger">
                    <strong>⚠️ Atenção:</strong> {pior['Colaborador(a)']}<br>
                    <span style="font-size: 14px; color: #dc3545;"><b>{pior['Percentual']:.1f}%</b></span> ({int(pior['Horas_Realizadas'])}h/{int(pior['horas totais'])}h)
                </div>
                <div class="highlight-box warning">
                    <strong>📚 Cursos:</strong><br>
                    ✅ {cursos_concluidos} | 🔄 {cursos_andamento} | ❌ {cursos_pendentes}
                </div>
            </div>
        </div>

        <div class="section">
            <div class="section-title">📈 PROGRESSO POR COLABORADOR (Horas Planejadas vs Realizadas)</div>
            <div class="chart-container">
    """
    
    # Gráfico de barras - Progresso
    max_horas = df_merged['horas totais'].max()
    for _, row in df_merged.sort_values('Percentual', ascending=False).iterrows():
        color = '#28a745' if row['Percentual'] >= 70 else ('#ff9800' if row['Percentual'] >= 30 else '#dc3545')
        width_total = (row['horas totais'] / max_horas * 100)
        width_realizado = (row['Horas_Realizadas'] / max_horas * 100)
        
        html_content += f"""
                <div class="bar-row">
                    <div class="bar-label">{row['Colaborador(a)'][:20]}</div>
                    <div class="bar-container">
                        <div class="bar-fill" style="width: {width_realizado}%; background: {color};">
                            <span class="bar-text">{row['Percentual']:.0f}%</span>
                        </div>
                    </div>
                    <div class="bar-value" style="color: {color};">{int(row['Horas_Realizadas'])}h / {int(row['horas totais'])}h</div>
                </div>
        """
    
    html_content += """
            </div>
        </div>

        <div class="section">
            <div class="section-title">⏱️ RITMO NECESSÁRIO PARA CUMPRIR O PRAZO (Horas por dia útil)</div>
            <div class="chart-container" style="position: relative;">
    """
    
    # Gráfico de ritmo
    max_ritmo = max(df_pace['Ritmo_Necessario'].max(), 3)
    for _, row in df_pace.iterrows():
        ritmo = row['Ritmo_Necessario']
        if ritmo <= 0:
            color = '#28a745'
            status = '✅'
        elif ritmo <= 1:
            color = '#2ecc71'
            status = '🟢'
        elif ritmo <= 1.5:
            color = '#f1c40f'
            status = '🟡'
        elif ritmo <= 2:
            color = '#e67e22'
            status = '🟠'
        else:
            color = '#e74c3c'
            status = '🔴'
        
        width = min((ritmo / max_ritmo * 100), 100) if ritmo > 0 else 0
        
        html_content += f"""
                <div class="ritmo-bar">
                    <div class="ritmo-name">{row['Colaborador(a)'][:18]}</div>
                    <div class="ritmo-container">
                        <div class="ritmo-fill" style="width: {width}%; background: {color};"></div>
                        <div class="ritmo-line" style="left: {1/max_ritmo*100}%; background: #2ecc71;"></div>
                        <div class="ritmo-line" style="left: {2/max_ritmo*100}%; background: #e74c3c;"></div>
                    </div>
                    <div class="ritmo-value" style="color: {color};">{status} {ritmo:.1f}h/dia</div>
                </div>
        """
    
    html_content += f"""
            </div>
            <div style="font-size: 8px; color: #666; margin-top: 5px;">
                Legenda: 🔵 Tranquilo (≤1h) | 🟢 Bom Ritmo (1-1.5h) | 🟡 Atenção (1.5-2h) | 🟠 Crítico (2-3h) | 🔴 Plano de Ação (>3h)
            </div>
        </div>

        <div class="page-break"></div>
        
        <div class="section">
            <div class="section-title">📋 DETALHAMENTO POR COLABORADOR</div>
    """
    
    # Detalhamento compacto
    for i, (_, row) in enumerate(df_merged.sort_values('Percentual', ascending=False).iterrows()):
        if i > 0 and i % 4 == 0:
            html_content += '<div class="page-break"></div>'
        
        df_colab = df_real[df_real['Colaborador(a)'] == row['Colaborador(a)']].copy()
        
        concluidos = len(df_colab[df_colab['Status'] == 'Concluído'])
        andamento = len(df_colab[df_colab['Status'] == 'Em Andamento'])
        pendentes = len(df_colab[df_colab['Status'] == 'Pendente'])
        
        ritmo_colab = df_pace[df_pace['Colaborador(a)'] == row['Colaborador(a)']]['Ritmo_Necessario'].values[0]
        color = '#28a745' if row['Percentual'] >= 70 else ('#ff9800' if row['Percentual'] >= 30 else '#dc3545')
        # Cores: Azul (≤1h), Verde (1-1.5h), Amarelo (1.5-2h), Laranja (2-3h), Vermelho (>3h)
        ritmo_color = '#3498db' if ritmo_colab <= 1 else ('#2ecc71' if ritmo_colab <= 1.5 else ('#f1c40f' if ritmo_colab <= 2 else ('#e67e22' if ritmo_colab <= 3 else '#e74c3c')))
        
        html_content += f"""
            <div class="colaborador-section">
                <div class="colaborador-header">
                    <span class="colaborador-name">👤 {row['Colaborador(a)']}</span>
                    <span style="color: {color}; font-weight: bold; font-size: 12px;">{row['Percentual']:.1f}%</span>
                </div>
                <div class="progress-bar" style="margin-bottom: 6px; height: 10px;">
                    <div class="progress-fill" style="width: {min(row['Percentual'], 100)}%; background: {color};"></div>
                </div>
                <div class="info-grid">
                    <div class="info-item">
                        <div class="info-value" style="color: #1976d2;">{int(row['horas totais'])}h</div>
                        <div class="info-label">Planejado</div>
                    </div>
                    <div class="info-item">
                        <div class="info-value" style="color: #28a745;">{int(row['Horas_Realizadas'])}h</div>
                        <div class="info-label">Concluído</div>
                    </div>
                    <div class="info-item">
                        <div class="info-value" style="color: #dc3545;">{int(row['Horas_Pendentes'])}h</div>
                        <div class="info-label">Pendente</div>
                    </div>
                    <div class="info-item">
                        <div class="info-value" style="color: {ritmo_color};">{ritmo_colab:.1f}h</div>
                        <div class="info-label">Ritmo/dia</div>
                    </div>
                </div>
                <div style="font-size: 9px; margin-bottom: 4px;">
                    <span class="status-green">✅ {concluidos}</span> |
                    <span class="status-yellow">🔄 {andamento}</span> |
                    <span class="status-red">❌ {pendentes}</span>
                </div>
                <table>
                    <tr><th>Curso</th><th style="width: 45px;">Carga</th><th style="width: 70px;">Status</th></tr>
        """
        
        for _, curso in df_colab.iterrows():
            icon = '✅' if curso['Status'] == 'Concluído' else ('🔄' if curso['Status'] == 'Em Andamento' else '❌')
            status_class = 'status-green' if curso['Status'] == 'Concluído' else ('status-yellow' if curso['Status'] == 'Em Andamento' else 'status-red')
            
            html_content += f"""
                    <tr>
                        <td>{str(curso['Curso'])[:50]}{'...' if len(str(curso['Curso'])) > 50 else ''}</td>
                        <td>{int(curso['Carga Horária'])}h</td>
                        <td class="{status_class}">{icon}</td>
                    </tr>
            """
        
        html_content += """
                </table>
            </div>
        """
    
    html_content += """
        </div>
    </body>
    </html>
    """
    
    return html_content


# ==================== INTERFACE PRINCIPAL ====================

def main():
    st.markdown('<h1 class="main-header">📚 Relatório de Cursos</h1>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Automatização de relatórios de acompanhamento de capacitação</p>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/training.png", width=80)
        st.markdown("### 📁 Upload de Dados")
        
        uploaded_file = st.file_uploader(
            "Selecione o arquivo Excel",
            type=['xlsx', 'xls'],
            help="O arquivo deve conter as abas 'Plano' e 'Real/Realizado'"
        )
        
        st.markdown("---")
        st.markdown("### ℹ️ Instruções")
        st.markdown("""
        1. Faça upload do arquivo Excel
        2. Visualize o dashboard
        3. Clique em **Gerar PDF** para exportar
        """)
    
    # Aguarda upload do arquivo
    if uploaded_file is None:
        st.info("👆 Por favor, faça upload do arquivo Excel na barra lateral.")
        st.markdown("""
        ### 📋 Formato esperado do arquivo:
        
        O arquivo Excel deve conter **duas abas**:
        
        **1. Aba "Plano"** - Horas planejadas por colaborador:
        | Id colaborador(a) | Colaborador(a) | horas totais |
        |-------------------|----------------|--------------|
        | 123456 | Nome do Colaborador | 200 |
        
        **2. Aba "Real"** - Cursos realizados:
        | Id colaborador(a) | Colaborador(a) | Curso | Carga Horária | Finalizou o curso? |
        |-------------------|----------------|-------|---------------|-------------------|
        | 123456 | Nome do Colaborador | Nome do Curso | 10 | Sim/Não |
        """)
        st.stop()
    else:
        df_plano, df_real = load_data(uploaded_file)
        st.sidebar.success("✅ Arquivo carregado com sucesso!")
    
    # Processa dados
    df_merged, df_real = process_data(df_plano, df_real)
    
    # ==================== PÁGINA 1: STORYTELLING + RESUMO ====================
    
    # Storytelling
    st.markdown("---")
    st.markdown("## 🎯 Contexto e Objetivo")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        <div class="storytelling-box">
        <h4>📖 Sobre este Plano de Desenvolvimento</h4>
        <p>Este relatório apresenta o acompanhamento do <strong>Plano de Capacitação da Equipe</strong>, 
        com foco no desenvolvimento de competências em:</p>
        <ul>
            <li>🎯 <strong>Liderança</strong> - Gestão de equipes e comunicação estratégica</li>
            <li>📊 <strong>Estatística e Dados</strong> - Análise e interpretação de informações</li>
            <li>💻 <strong>Ferramentas Digitais</strong> - Power BI, Python, SQL e automação</li>
        </ul>
        <p>O objetivo é desenvolver a equipe para melhorar a <strong>tomada de decisão baseada em dados</strong> 
        e aumentar a <strong>eficiência operacional</strong>.</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        # Métricas rápidas
        total_planejado = df_merged['horas totais'].sum()
        total_realizado = df_merged['Horas_Realizadas'].sum()
        percentual_geral = (total_realizado / total_planejado * 100)
        
        st.metric("👥 Colaboradores", len(df_merged))
        st.metric("📚 Total de Cursos", len(df_real))
        st.metric("⏱️ Horas Planejadas", f"{int(total_planejado)}h")
    
    # ==================== RESUMO EXECUTIVO ====================
    
    st.markdown("---")
    st.markdown("## 📊 Resumo Executivo")
    
    # KPIs em cards
    col1, col2, col3, col4 = st.columns(4)
    
    cursos_concluidos = len(df_real[df_real['Status'] == 'Concluído'])
    cursos_andamento = len(df_real[df_real['Status'] == 'Em Andamento'])
    cursos_pendentes = len(df_real[df_real['Status'] == 'Pendente'])
    
    with col1:
        st.metric(
            "🎯 Progresso Geral",
            f"{percentual_geral:.1f}%",
            delta=f"{total_realizado:.0f}h concluídas"
        )
    
    with col2:
        st.metric(
            "✅ Cursos Concluídos",
            cursos_concluidos,
            delta=f"{(cursos_concluidos/len(df_real)*100):.1f}%"
        )
    
    with col3:
        st.metric(
            "🔄 Em Andamento",
            cursos_andamento,
            delta=None
        )
    
    with col4:
        st.metric(
            "❌ Pendentes",
            cursos_pendentes,
            delta=f"-{(cursos_pendentes/len(df_real)*100):.1f}%",
            delta_color="inverse"
        )
    
    # Gráficos lado a lado
    col1, col2 = st.columns([3, 2])
    
    with col1:
        fig_bar = create_bar_chart(df_merged)
        st.plotly_chart(fig_bar, use_container_width=True, key="bar_chart")
    
    with col2:
        fig_pie, percentual_geral, total_realizado, total_planejado = create_pie_chart(df_merged)
        st.plotly_chart(fig_pie, use_container_width=True, key="pie_chart")
    
    # ==================== GRÁFICO DE RITMO ====================
    
    st.markdown("---")
    st.markdown("## ⏱️ Análise de Ritmo para Cumprimento do Prazo")
    
    fig_pace, df_pace, dias_totais, dias_estudo, dias_uteis_total = create_pace_chart(df_merged)
    
    # Info box explicativo
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("📅 Prazo Final", "20/12/2026")
    with col2:
        st.metric("⏳ Dias Corridos", f"{dias_totais}")
    with col3:
        st.metric("📆 Dias Úteis", f"{dias_uteis_total}")
    with col4:
        st.metric("📚 Dias Efetivos (70%)", f"{dias_estudo}")
    with col5:
        plano_acao = len(df_pace[df_pace['Ritmo_Necessario'] > 3])
        st.metric("🔴 Plano de Ação", f"{plano_acao} pessoas", delta=None if plano_acao == 0 else "atenção", delta_color="inverse")
    
    st.plotly_chart(fig_pace, use_container_width=True, key="pace_chart")
    
    # Legenda explicativa
    st.markdown(f"""
    <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; margin-top: -20px;">
        <strong>📊 Como interpretar:</strong>
        <span style="color: #3498db;">●</span> <b>Tranquilo</b> (≤1h/dia) |
        <span style="color: #2ecc71;">●</span> <b>Bom Ritmo</b> (1-1.5h/dia) |
        <span style="color: #f1c40f;">●</span> <b>Atenção</b> (1.5-2h/dia) |
        <span style="color: #e67e22;">●</span> <b>Crítico</b> (2-3h/dia) |
        <span style="color: #e74c3c;">●</span> <b>Plano de Ação</b> (>3h/dia)
        <br><small>💡 <b>Considerando 70% dos dias úteis</b> (margem para reuniões e imprevistos)</small>
    </div>
    """, unsafe_allow_html=True)
    
    # Destaques
    st.markdown("### 🏆 Destaques")
    
    melhor = df_merged.loc[df_merged['Percentual'].idxmax()]
    pior = df_merged.loc[df_merged['Percentual'].idxmin()]
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"""
        <div class="success-box">
        <h4>🥇 Maior Progresso</h4>
        <p><strong>{melhor['Colaborador(a)']}</strong></p>
        <p style="font-size: 24px; color: #28a745;"><strong>{melhor['Percentual']:.1f}%</strong> concluído</p>
        <p>{int(melhor['Horas_Realizadas'])}h de {int(melhor['horas totais'])}h planejadas</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="danger-box">
        <h4>⚠️ Requer Atenção</h4>
        <p><strong>{pior['Colaborador(a)']}</strong></p>
        <p style="font-size: 24px; color: #dc3545;"><strong>{pior['Percentual']:.1f}%</strong> concluído</p>
        <p>{int(pior['Horas_Realizadas'])}h de {int(pior['horas totais'])}h planejadas</p>
        </div>
        """, unsafe_allow_html=True)
    
    # ==================== DETALHAMENTO ====================
    
    st.markdown("---")
    st.markdown("## 📋 Detalhamento por Colaborador")
    
    # Seletor de colaborador
    colaborador_selecionado = st.selectbox(
        "Selecione um colaborador para ver detalhes:",
        options=df_merged.sort_values('Percentual', ascending=False)['Colaborador(a)'].tolist()
    )
    
    # Dados do colaborador selecionado
    dados_colab = df_merged[df_merged['Colaborador(a)'] == colaborador_selecionado].iloc[0]
    df_cursos_colab = df_real[df_real['Colaborador(a)'] == colaborador_selecionado].copy()
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        # Gauge de progresso
        fig_gauge = create_gauge_chart(dados_colab['Percentual'], colaborador_selecionado)
        st.plotly_chart(fig_gauge, use_container_width=True)
        
        # Métricas
        st.markdown(f"**Horas Planejadas:** {int(dados_colab['horas totais'])}h")
        st.markdown(f"**Horas Concluídas:** {int(dados_colab['Horas_Realizadas'])}h")
        st.markdown(f"**Horas Pendentes:** {int(dados_colab['Horas_Pendentes'])}h")
    
    with col2:
        # Contagem por status
        status_counts = df_cursos_colab['Status'].value_counts()
        
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.metric("✅ Concluídos", status_counts.get('Concluído', 0))
        with col_b:
            st.metric("🔄 Em Andamento", status_counts.get('Em Andamento', 0))
        with col_c:
            st.metric("❌ Pendentes", status_counts.get('Pendente', 0))
        
        # Tabela de cursos
        st.markdown("#### Cursos")
        
        df_display = df_cursos_colab[['Curso', 'Carga Horária', 'Status']].copy()
        df_display['Ícone'] = df_display['Status'].apply(get_status_icon)
        df_display = df_display[['Ícone', 'Curso', 'Carga Horária', 'Status']]
        
        # Aplica estilo
        def highlight_status(row):
            if row['Status'] == 'Concluído':
                return ['background-color: #d4edda'] * len(row)
            elif row['Status'] == 'Em Andamento':
                return ['background-color: #fff3cd'] * len(row)
            else:
                return ['background-color: #f8d7da'] * len(row)
        
        st.dataframe(
            df_display.style.apply(highlight_status, axis=1),
            use_container_width=True,
            height=300
        )
    
    # ==================== VISÃO GERAL DE TODOS ====================
    
    st.markdown("---")
    st.markdown("## 👥 Visão Geral - Todos os Colaboradores")
    
    # Expanders para cada colaborador
    for _, row in df_merged.sort_values('Percentual', ascending=False).iterrows():
        df_colab = df_real[df_real['Colaborador(a)'] == row['Colaborador(a)']].copy()
        
        status_counts = df_colab['Status'].value_counts()
        concluidos = status_counts.get('Concluído', 0)
        andamento = status_counts.get('Em Andamento', 0)
        pendentes = status_counts.get('Pendente', 0)
        
        # Ícone baseado no progresso
        if row['Percentual'] >= 70:
            icon = "🟢"
        elif row['Percentual'] >= 30:
            icon = "🟡"
        else:
            icon = "🔴"
        
        with st.expander(f"{icon} **{row['Colaborador(a)']}** - {row['Percentual']:.1f}% ({int(row['Horas_Realizadas'])}h / {int(row['horas totais'])}h)"):
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Progresso", f"{row['Percentual']:.1f}%")
            with col2:
                st.metric("✅ Concluídos", concluidos)
            with col3:
                st.metric("🔄 Andamento", andamento)
            with col4:
                st.metric("❌ Pendentes", pendentes)
            
            # Barra de progresso visual
            progress_color = '#28a745' if row['Percentual'] >= 70 else ('#ffc107' if row['Percentual'] >= 30 else '#dc3545')
            st.markdown(f"""
            <div style="background: #e0e0e0; border-radius: 10px; height: 20px; overflow: hidden;">
                <div style="background: {progress_color}; height: 100%; width: {min(row['Percentual'], 100)}%; border-radius: 10px;"></div>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("")
            
            # Lista de cursos resumida
            df_display = df_colab[['Curso', 'Carga Horária', 'Status']].copy()
            df_display['Ícone'] = df_display['Status'].apply(get_status_icon)
            
            st.dataframe(
                df_display[['Ícone', 'Curso', 'Carga Horária', 'Status']],
                use_container_width=True,
                hide_index=True
            )
    
    # ==================== BOTÃO GERAR PDF ====================
    
    st.markdown("---")
    st.markdown("## 📄 Exportar Relatório")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
        <div style="text-align: center; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 15px; color: white;">
            <h3>📥 Gerar Relatório em PDF</h3>
            <p>Clique no botão abaixo para gerar o relatório executivo em PDF (máx. 3 páginas)</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("")
        
        if st.button("📄 Gerar PDF", type="primary", use_container_width=True):
            with st.spinner("Gerando PDF..."):
                html_content = generate_pdf_content(
                    df_merged, df_real, percentual_geral, 
                    total_realizado, total_planejado
                )
                
                # Salva HTML
                st.download_button(
                    label="📥 Baixar HTML do Relatório",
                    data=html_content,
                    file_name=f"relatorio_cursos_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                    mime="text/html",
                    use_container_width=True
                )
                
                st.success("✅ Relatório gerado! Abra o arquivo HTML no navegador e use Ctrl+P para salvar como PDF.")
                st.info("💡 **Dica:** No Chrome/Edge, ao imprimir, selecione 'Salvar como PDF' e marque 'Gráficos de fundo' nas opções.")


if __name__ == "__main__":
    main()
