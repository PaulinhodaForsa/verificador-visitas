# Verificador de Visitas

Esta ferramenta analisa um arquivo de chat do WhatsApp exportado em TXT e cruza as mensagens com uma planilha Excel contendo as lojas. Ao final é gerada uma nova planilha assinalando os dias em que cada loja foi citada.

## Pré-requisitos

- Python 3.10+
- Pacotes `pandas`, `rapidfuzz`, `unidecode`, `openpyxl` e `tkinter`.

Instale os pacotes necessários com:

```bash
pip install pandas rapidfuzz unidecode openpyxl
```

## Como usar

Execute o script:

```bash
python verificador_visitas_v3.py
```

Uma janela será aberta pedindo dois arquivos:

1. **TXT do WhatsApp** – exporte a conversa usando *Exportar conversa* no WhatsApp. Cada linha deve seguir o formato:

   ```
   [08:15, 26/05/2025] Rodrigo: Visitei loja Carapicuiba
   ```

2. **Planilha base** – arquivo Excel com uma coluna chamada `Loja` e colunas para cada dia da semana (`Segunda`, `Terça`, `Quarta`, `Quinta`, `Sexta`, `Sábado`). Exemplo:

   | Loja        | Segunda | Terça | Quarta | Quinta | Sexta | Sábado |
   | ----------- | ------- | ----- | ------ | ------ | ----- | ------ |
   | Carapicuiba |         |       |        |        |       |        |
   | Ferraz      |         |       |        |        |       |        |

Clique em **Gerar Relatório** e o programa criará `relatorio_<nome-do-txt>.xlsx` na mesma pasta.

## Entradas esperadas

- **TXT**: arquivo de chat do WhatsApp no formato `[HH:MM, dd/mm/yyyy] Nome: mensagem`.
- **Excel**: planilha com a lista de lojas e colunas dos dias da semana como visto acima. As células vazias serão preenchidas com `X` quando uma visita for encontrada.

Ajuste a constante `data_inicio` dentro de `verificador_visitas_v3.py` para refletir a segunda-feira da semana presente nas mensagens do WhatsApp.
