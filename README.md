# Stock Management Automation

Este repositório contém um script em Python para processar e organizar dados de estoque provenientes de arquivos Excel. Ele realiza limpeza de dados, aplica mapeamentos e gera relatórios consolidados no formato Excel, incluindo formatação como tabelas dinâmicas.

---

## Funcionalidades

- **Processamento de Arquivos Excel:**
  - Lê e processa arquivos Excel localizados em um diretório de entrada.
- **Mapeamento de Códigos:**
  - Aplica mapeamentos personalizados para códigos de modelo e revendas.
- **Atualização de Dados:**
  - Calcula o tempo de estoque baseado em número de dias.
  - Diferencia opcionais de modelos com base em datas de corte.
- **Geração de Relatórios:**
  - Cria um arquivo Excel consolidado com os dados processados.
  - Aplica formatação de tabela dinâmica no Excel.

---

## Requisitos

- **Python 3.8 ou superior.**
- **Bibliotecas necessárias:**
  - pandas
  - openpyxl

Para instalar as dependências, execute:
```bash
pip install pandas openpyxl
```

---

## Configuração

1. Certifique-se de que os arquivos Excel estejam localizados no diretório configurado como entrada no script:
   ```python
   PASTA_INPUT = "C:\\data\\input"
   ```

2. Configure o diretório de saída para os relatórios gerados:
   ```python
   PASTA_OUTPUT = "C:\\data\\output"
   ```

3. Certifique-se de que os mapeamentos personalizados (como `model_mapping` e `revenda_mapping`) estejam ajustados conforme suas necessidades.

---

## Como Executar

1. Clone o repositório para sua máquina local:
   ```bash
   git clone https://github.com/seu-usuario/stock-management-automation.git
   cd stock-management-automation
   ```

2. Execute o script principal:
   ```bash
   python main.py
   ```

3. Após a execução, o relatório consolidado estará disponível no diretório configurado em `PASTA_OUTPUT`.

---

## Estrutura do Relatório

### Colunas Calculadas

- **Options:** Mapeamento baseado nos códigos de modelo.
- **TIME:** Classificação do tempo de estoque em intervalos (e.g., 0-30, 31-60 dias).
- **Dealership_Name:** Substitui nomes de revendas com base no mapeamento configurado.

### Formatação como Tabela

- O relatório final é formatado como tabela Excel para facilitar a análise.

---

## Observações Importantes

- O script ignora arquivos temporários ou corrompidos que começam com `~$`.
- Certifique-se de ajustar as datas de corte e mapeamentos conforme suas regras de negócio.

---

## Contribuições

Contribuições são bem-vindas! Para relatar problemas ou sugerir melhorias, envie uma solicitação via [Issues](https://github.com/seu-usuario/stock-management-automation/issues) ou abra um pull request.

---

## Licença

Este projeto está licenciado sob a MIT License.

---

## Autor

- **Gabriel Matuck**
- **E-mail:** [gabriel.matuck1@gmail.com](mailto:gabriel.matuck1@gmail.com)

---

Automatize o gerenciamento de estoque com eficiência e precisão!

