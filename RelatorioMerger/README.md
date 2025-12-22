# 📊 Relatório Merger - Virtron

Ferramenta interna em Python para mesclar e formatar relatórios da **Virtron Energia Solar** (Pós-Venda e VendasI). Criada para agilizar o trabalho do setor de TI.

---

## ✨ Funcionalidades

- Mescla arquivos Excel (**Pós-Venda** e **VendasI**) pelo campo **Título**.  
- Remove colunas desnecessárias automaticamente.  
- Renomeia colunas para padronização.  
- Formata o Excel final com:
  - Cabeçalho verde
  - Bordas finas em todas as células
  - Centralização do conteúdo a partir da coluna 4
  - Ajuste automático da largura das colunas

---

## 🛠 Requisitos

- Python 3.10 ou superior  
- Bibliotecas:
  - `pandas`
  - `openpyxl`
  - `customtkinter`
  - `tkinter` (já incluso no Python)

Instalação das bibliotecas:
pip install pandas openpyxl customtkinter

## 🚀 Como usar

1. Abra a pasta dist
2. Execute o arquivo `RelatorioMerger.exe`.  
3. Selecione o arquivo de **Pós-Venda**.  
4. Selecione o arquivo de **VendasI**.  
5. Clique em **Iniciar Mesclagem**.  
6. O arquivo mesclado será salvo na mesma pasta do arquivo de VendasI como `RelatórioMesclado.xlsx`.  

> ⚠️ Observação: Esta ferramenta foi desenvolvida especificamente para os relatórios da Virtron e pode não funcionar corretamente com outros formatos.

---

## 📝 Sobre

Criado por **Sergio Barbosa**, para uso interno do setor de TI da Virtron, com o objetivo de agilizar a mesclagem e formatação de relatórios Excel.

---

## 📁 Distribuição

- Para enviar para alguém, inclua o **arquivo executável** (`.exe`) e quaisquer ícones ou recursos adicionais utilizados.  
- A pessoa não precisa instalar Python ou bibliotecas para usar o `.exe`.
