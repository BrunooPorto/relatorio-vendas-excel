# 📊 Automação de Relatório de Vendas em Excel

Script em **Python** que transforma uma base de vendas (CSV) em um relatório
**Excel profissional com 4 abas formatadas** — pronto para enviar a um gestor
ou cliente, sem trabalho manual.

> Ideal para empresas que ainda montam relatórios "na mão" e perdem horas
> copiando, colando e formatando planilhas toda semana.

---

## ✨ O que ele entrega

O relatório gerado (`relatorio_vendas.xlsx`) contém:

| Aba | Conteúdo |
|-----|----------|
| **Resumo** | KPIs do período: faturamento total, nº de pedidos, itens vendidos, ticket médio, categoria e vendedor destaque |
| **Por Categoria** | Faturamento, itens, pedidos, ticket médio e % do total por categoria |
| **Por Vendedor** | Mesmo recorte, por vendedor — para acompanhar metas |
| **Dados Brutos** | A base completa, com cabeçalho congelado e filtros prontos |

Tudo com **cabeçalhos coloridos, formato de moeda (R$), linhas alternadas,
bordas, larguras automáticas, filtros e painéis congelados**.

---

## 🚀 Como usar

```bash
# 1. Instale as dependências
pip install pandas openpyxl

# 2. Rode o script
python relatorio_vendas_auto.py
```

Se você **não** tiver um arquivo de dados, o script gera automaticamente uma
base de exemplo (`vendas.csv`) com 300 vendas fictícias — assim você vê o
resultado na hora.

### Usando seus próprios dados

Basta ter um CSV com as colunas:

```
id_venda, data, produto, categoria, vendedor, regiao, quantidade, preco_unitario, valor_total
```

E rodar apontando para ele:

```bash
python relatorio_vendas_auto.py --entrada minhas_vendas.csv --saida relatorio.xlsx
```

---

## 🛠️ Tecnologias

- **Python 3.10+**
- **pandas** — leitura, agregação e cálculo dos indicadores
- **openpyxl** — geração e formatação avançada do Excel

---

## 💡 Sobre

Projeto de portfólio focado em **automação de tarefas repetitivas de escritório**.
Posso adaptar o relatório ao seu modelo de planilha, adicionar gráficos,
enviar por e-mail automaticamente ou agendar a geração diária/semanal.

📬 **Vamos conversar sobre como automatizar seus relatórios?**
