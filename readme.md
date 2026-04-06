# 📊 Sistema de Relatórios de Atendimentos

---

Projeto desenvolvido em Python com foco em análise, automação e visualização de atendimentos.
O sistema permite processar múltiplos arquivos, gerar métricas e exportar relatórios profissionais.

## 🚀 Funcionalidades

* Upload de múltiplos arquivos (CSV e Excel)
* Processamento e limpeza automática dos dados
* Cálculo de métricas:

  * Total de atendimentos
  * Atendimentos por atendente
  * Tempo médio geral
  * Tempo médio por atendente
  * Atendimentos por dia
* Geração de relatórios:

  * TXT
  * Excel (com múltiplas abas)
* Interface web com Django
* Download de relatório direto pelo navegador
* Validação inteligente de arquivos (colunas obrigatórias)
* Tratamento automático de encoding

## 🛠 Tecnologias

* Python
* Pandas
* Django
* OpenPyXL
* Chardet

## 🚀 Roadmap 🗺️

v1.0 ✅

* Processamento de planilhas

v1.1 ✅

* Relatórios em TXT

v1.2 ✅

* Relatórios em Excel

v2.0 ✅

* Estrutura web com Django

v2.1 ✅

* Página HTML inicial

v2.2 ✅

* Upload de múltiplos arquivos
* Suporte a CSV e Excel (.xlsx)
* Validação de dados
* Normalização de nomes
* Métricas avançadas
* Download de relatório via navegador

## 🟡 Próximos passos 🟡

* Gráficos (dashboard visual)
* Melhor organização do código (services / utils)
* Feedback visual de sucesso no upload
* Melhor tratamento de erros na interface

## 🔮 Futuro

* Sistema de senhas para atendimento
* Registro automático de início e fim
* Dashboard em tempo real
* Banco de dados e histórico de relatórios
* Sistema de login e permissões

## ⚙️ Como rodar o projeto

cd relatorio_atendimentos
.\venv\Scripts\activate
cd webapp
python manage.py runserver

Acesse no navegador:
http://127.0.0.1:8000/