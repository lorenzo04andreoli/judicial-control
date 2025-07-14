# Gerador de Planilha Geral de Reeducandos

Este script em Python automatiza a criação de uma planilha geral com os dados de reeducandos que cumprem pena em juízo. Ele coleta informações individuais de diversos arquivos `.xlsx` com base em um modelo padrão de ficha e consolida tudo em uma única planilha.

---

## Funcionalidades

- Varre todos os arquivos `.xlsx` de uma pasta.
- Extrai automaticamente:
  - Nome
  - CPF
  - Telefone
  - Frequência de comparecimento (Mensal, Bimestral, Outros)
  - Nome do arquivo de origem
- Lida com variações de espaçamento e formatação nos campos de frequência (ex: `(X)`, `( X )`, `(  X)` etc).
- Gera uma planilha chamada `planilha_reeducandos.xlsx` com os dados unificados.

---

## 📂 Estrutura esperada dos arquivos `.xlsx`

Cada ficha individual deve conter:

| Célula | Conteúdo                    |
|--------|-----------------------------|
| B4     | Nome completo               |
| G5     | CPF                         |
| B8     | Telefone                    |
| A2     | Frequência com opções como:<br>`( X ) MENSAL ( ) BIMESTRAL ( ) OUTROS` |

---

## 🛠️ Requisitos

- Python 3.7+
- Bibliotecas:
  - `openpyxl`
  - `re` (já incluída na biblioteca padrão do Python)

Instalação da dependência:

```bash
pip install openpyxl
```

---

## 🚀 Como usar

1. Altere o caminho da variável `pasta_fichas` para apontar para a pasta que contém os arquivos `.xlsx`.

```python
pasta_fichas = r"C:\Caminho\Pasta"
```

2. Execute o script:

```bash
python script.py
```

3. O arquivo `planilha_reeducandos.xlsx` será criado na mesma pasta do script, contendo os dados de todos os reeducandos.

---

## 🧑‍💻 Autor

Criado por Lorenzo Andreoli (2025), para organização dos comparecimentos de reeducandos em juízo.

---

## 📌 Observação

Certifique-se de que todos os arquivos `.xlsx` sigam o padrão de estrutura indicado acima para garantir o correto funcionamento do script.