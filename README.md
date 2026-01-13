# 📊 Relatório de Acompanhamento de Cursos

Aplicação Streamlit para acompanhamento de progresso de cursos por colaboradores.

## 🚀 Funcionalidades

- Dashboard interativo com métricas de progresso
- Gráfico de ritmo necessário para cumprir prazo (20/12/2026)
- Detalhamento por colaborador
- Geração de relatório PDF executivo
- Cálculo de dias úteis (70% - margem para imprevistos)

## 📋 Pré-requisitos

- Python 3.10+
- Arquivo Excel com abas "Plano" e "Real/Realizado"

## 🛠️ Instalação Local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 🌐 Deploy no Streamlit Cloud

1. Conecte seu repositório GitHub ao [Streamlit Cloud](https://streamlit.io/cloud)
2. Selecione o repositório e o arquivo `app.py`
3. Deploy!

## 📁 Estrutura do Projeto

```
├── app.py                 # Aplicação principal
├── requirements.txt       # Dependências
├── .streamlit/
│   └── config.toml       # Configuração do tema
└── README.md
```

## 📊 Legenda do Ritmo

- 🔵 **Tranquilo** (≤1h/dia)
- 🟢 **Bom Ritmo** (1-1.5h/dia)
- 🟡 **Atenção** (1.5-2h/dia)
- 🟠 **Crítico** (2-3h/dia)
- 🔴 **Plano de Ação** (>3h/dia)
