# Projeto de Processamento de Planilhas - Colaboradores Sem Produção

## Descrição do Projeto

Este projeto foi desenvolvido com o objetivo de facilitar a análise e o processamento de planilhas contendo informações de colaboradores e suas produções. Ele foi feito **especificamente para mim**, com a intenção de automatizar a verificação de colaboradores sem produção ou com registros preenchidos de maneira errada.

A aplicação possui uma interface gráfica desenvolvida em **Tkinter**, que permite o processamento de dados extraídos de duas planilhas (`ps.xlsx` e `base_ps.xlsx`), e gera um relatório com as informações dos colaboradores que não registraram produção ou tiveram dados incorretos.

## Funcionalidades

1. **Processamento de Planilhas:**
   - O script lê as planilhas `ps.xlsx` e `base_ps.xlsx` e as processa para identificar:
     - Colaboradores que não geraram produção (não encontrados na planilha `base_ps`).
     - Registros errados ou inconsistentes, onde os dados da planilha `ps.xlsx` não coincidem com a `base_ps`.

2. **Interface Gráfica:**
   - A interface oferece a visualização de uma lista com todos os colaboradores sem produção ou com dados preenchidos incorretos.
   - Filtros de data e matrícula são aplicáveis para refinar a visualização.
   - Um botão para marcar registros como "Corrigidos" permite gerenciar rapidamente os registros já revisados.

3. **Geração de Relatório:**
   - Após o processamento das planilhas, é gerado um arquivo `colaboradores_sem_producao.xlsx` com a lista final dos colaboradores sem produção ou com preenchimento errado.

4. **Verificador de Produção (Função Adicional):**
   - A aplicação também oferece a possibilidade de abrir uma janela secundária (usando o comando "Abrir Verificador Produção") que permite revisar os dados detalhados de colaboradores sem produção.

## Funcionalidades do Verificador de Produção

- A janela "Verificador Produção" é acessada por um botão na interface principal e realiza os seguintes passos:
  - Verifica se há discrepâncias entre as planilhas `ps.xlsx` e `base_ps.xlsx`.
  - Exibe um relatório com os colaboradores que estão listados como "Sim" na coluna **"Foi para Campo Gerar Produção?"**, mas não possuem correspondência na planilha `base_ps`.
  - Permite a marcação dos registros como "Corrigido" diretamente na interface.

## Requisitos

- **Python 3.x**: O projeto foi desenvolvido com a versão mais recente do Python 3.
- **Bibliotecas**:
  - `pandas`: Para manipulação e processamento das planilhas Excel.
  - `tkinter`: Para criação da interface gráfica.
  - `openpyxl` (ou equivalente): Necessário para manipular arquivos `.xlsx` com o `pandas`.

Para instalar as dependências, execute o seguinte comando:

```bash
pip install pandas openpyxl
```

## Como Executar

1. **Prepare suas Planilhas**:
   - Coloque os arquivos `ps.xlsx` e `base_ps.xlsx` na mesma pasta onde o script `app.py` está localizado.
   - As planilhas devem ter o seguinte formato:

   **Planilha `ps.xlsx`:**

   | conc | conc2 | DATA       | MAT  | COLABORADOR          | PROCESSO | Foi para Campo Gerar Produção? | ... |
   |------|-------|------------|------|----------------------|----------|--------------------------------|-----|
   | ...  | ...   | 22/10/2024 | 10846| LEANDRO SOUZA ROSA   | CCM      | Sim                            | ... |

   **Planilha `base_ps.xlsx`:**

   | DATA       | NOME                           | MATRICULA | HR IN    | HR FN    | AL IN  | AL FN  | PROCESSO   | PLACA | PROJETO |
   |------------|--------------------------------|-----------|----------|----------|--------|--------|------------|-------|---------|
   | 22/10/2024 | MARCO ANTONIO DO NASCIMENTO VALADARES | 10849     | 7:30 AM  | 5:18 PM  | 12:00 PM | 1:00 PM | MANUTENÇÃO | LMJ3E79| 31124112103 |

2. **Execute o Script Principal**:
   - Abra o terminal ou prompt de comando na pasta do projeto.
   - Execute o script principal:

   ```bash
   python app.py
   ```

3. **Interaja com a Interface**:
   - Após rodar o script, a interface gráfica será exibida.
   - Use os filtros de **Data** e **Matrícula** para buscar dados específicos.
   - Você pode aplicar o filtro, reprocessar as planilhas ou abrir o verificador de produção.

## Personalização

Este projeto foi desenvolvido para o meu uso pessoal, com base nas planilhas que eu recebo periodicamente. Caso precise de ajustes para utilizar em outros contextos ou formatos de dados diferentes, é possível modificar o código conforme necessário. A estrutura de leitura e processamento das planilhas está localizada nas classes `ProcessadorPlanilha` nos dois arquivos.

## Licença

Este projeto foi feito **exclusivamente para uso pessoal**. Não há garantias de que ele funcionará em outros cenários, pois foi desenvolvido para atender a um caso específico. O código é de livre uso, mas não é garantido que ele se encaixe em outros requisitos sem ajustes.

---

**Nota:** Este projeto foi desenvolvido por mim para automatizar um processo específico em meu trabalho. A estrutura e os detalhes estão adaptados para meu fluxo de trabalho, portanto, ajustes podem ser necessários para usá-lo em outro contexto.
