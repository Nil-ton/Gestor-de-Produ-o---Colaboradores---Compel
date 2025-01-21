# Gestor de Produção de Colaboradores

Este repositório contém dois scripts desenvolvidos especificamente para **facilitar a verificação de dados de produção** dos colaboradores, com o objetivo de **reduzir a verificação manual** e otimizar o processo de conferência das planilhas de produção.

Os scripts foram criados para uso pessoal, considerando que eu (usuário) utilizo uma planilha de controle (a "base_ps.xlsx") e outra planilha externa (a "ps.xlsx") que contém informações adicionais sobre a produção. O sistema foi desenvolvido para **automatizar a identificação de erros** e colaboradores sem produção registrada, diminuindo a necessidade de verificar manualmente as informações nos papéis.

## Funcionalidades

- **Processamento de Planilhas**: O sistema lê e processa as planilhas `base_ps.xlsx` e `ps.xlsx`, identificando rapidamente os colaboradores que não registraram produção ou preencheram os dados de forma incorreta.
  
- **Interface Gráfica**: Uma interface gráfica foi criada para visualizar os dados processados, com filtros para pesquisa por data e matrícula, e a possibilidade de marcar os registros como corrigidos diretamente no sistema.

- **Geração de Relatório**: Após o processamento, é gerado um arquivo Excel com os colaboradores sem produção ou com registros errados, ajudando a manter o controle e a correção dos dados de forma eficiente.

## Contexto de Criação

Este sistema foi desenvolvido **para meu uso pessoal**, com o objetivo de:

1. **Facilitar a verificação de produção** de colaboradores, já que tradicionalmente eu realizava o controle manualmente em planilhas.
2. **Automatizar a comparação** entre os dados que preencho na minha planilha (`base_ps.xlsx`) e os dados externos provenientes de outra planilha (`ps.xlsx`).
3. **Reduzir o tempo de verificação manual**, que, sem uma ferramenta automatizada, seria feito de maneira mais demorada e propensa a erros.

A ideia é melhorar a eficiência do meu trabalho, economizando tempo e garantindo maior precisão nas verificações.

## Estrutura do Repositório

O repositório contém os seguintes scripts:

1. **gestor_producao_colaboradores.py**:
    - Script que processa os dados das planilhas, identificando os colaboradores sem produção registrada ou com dados errados, gerando um arquivo Excel com as informações de forma consolidada.
    - Utiliza a biblioteca `pandas` para manipulação dos dados e `tkinter` para a criação da interface gráfica.

2. **visualizar_producao_colaboradores.py**:
    - Script que oferece uma interface gráfica para visualizar os dados processados.
    - Permite filtrar as informações por data e matrícula, além de permitir a marcação dos registros como corrigidos.

## Requisitos

Para executar os scripts, é necessário ter o Python instalado, juntamente com as seguintes bibliotecas:

- `pandas`
- `tkinter`
- `openpyxl` (para ler/escrever arquivos Excel)

Instale as dependências necessárias com o seguinte comando:

```bash
pip install pandas openpyxl
```

**Nota**: O `tkinter` geralmente já vem instalado com o Python, mas caso contrário, você pode instalá-lo com o comando:

```bash
pip install tk
```

## Como Usar

### 1. Preparação das Planilhas

O script requer duas planilhas em formato Excel (`.xlsx`):

- **base_ps.xlsx**: Planilha onde preencho os dados de colaboradores e produção.
- **ps.xlsx**: Planilha externa com informações adicionais de produção, indicando se o colaborador foi para o campo gerar produção.

### 2. Executando os Scripts

1. **Executar o Script de Processamento (gestor_producao_colaboradores.py)**

   Para rodar o script que processa as planilhas e gera o arquivo `colaboradores_sem_producao.xlsx`, basta executar o seguinte comando:

   ```bash
   python gestor_producao_colaboradores.py
   ```

   O script irá processar as planilhas e gerar um arquivo Excel com as informações consolidadas.

2. **Executar o Script da Interface Gráfica (visualizar_producao_colaboradores.py)**

   Para abrir a interface gráfica que permite visualizar e interagir com os dados, execute o seguinte comando:

   ```bash
   python visualizar_producao_colaboradores.py
   ```

   A interface permite visualizar os dados processados, aplicar filtros por data e matrícula, e marcar os registros como corrigidos.

### 3. Funcionalidades na Interface Gráfica

- **Aplicar Filtro**: Permite filtrar os dados por data (formato dd/mm/yyyy) e matrícula.
- **Marcar como Corrigido**: Permite marcar um registro como corrigido diretamente na interface.
- **Reprocessar Planilhas**: Reprocessa as planilhas e atualiza os dados na interface gráfica.

## Exemplo de Uso

- Com a interface gráfica, você pode visualizar facilmente os colaboradores que não registraram produção ou preencheram os dados incorretamente, aplicar filtros para pesquisar por data ou matrícula e corrigir diretamente os registros.
- O sistema também gera o relatório `colaboradores_sem_producao.xlsx`, que pode ser utilizado para corrigir ou ajustar os dados manualmente, caso necessário.

## Contribuições

Este projeto foi feito para meu uso pessoal, mas se alguém quiser contribuir ou sugerir melhorias, fique à vontade para abrir uma *issue* ou enviar um *pull request*.

## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para mais detalhes.
