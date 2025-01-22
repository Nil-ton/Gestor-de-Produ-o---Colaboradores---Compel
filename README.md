# Sistema de Gestão de Produção e Registros de Colaboradores

Este repositório contém um sistema desenvolvido com o objetivo de **facilitar o meu trabalho** no gerenciamento e processamento dos dados relacionados à produção e ao acompanhamento de colaboradores em campo. O programa foi criado para atender a necessidades específicas do meu dia a dia, simplificando o processo de organização e análise de informações.

O sistema utiliza duas planilhas principais: **ps.xlsx** e **base_ps.xlsx**, que contêm dados distintos usados no acompanhamento das atividades dos colaboradores e produção no campo.

## Estrutura das Planilhas

### 1. `ps.xlsx` (Planilha de Produção)

Esta planilha contém informações sobre as atividades dos colaboradores em um dia específico. Sua estrutura é composta pelas seguintes colunas:

| **conc**        | **conc2**      | **DATA**      | **MAT** | **COLABORADOR**         | **PROCESSO** | **Foi para Campo Gerar Produção?** | **Participou do DDS?** | **MAT ENC** | **ENCARREGADO**         | **PLACA** | **PLACA INCORRETA?** | **PREFIXO** | **STATUS** | **JUSTIFICATIVA** | **MÊS** | **processo** | **const base veic** | **BASE** | **Log 1** | **Log 2** | **Log 3** |
|-----------------|----------------|---------------|--------|-------------------------|--------------|-----------------------------------|-------------------------|-------------|-------------------------|-----------|----------------------|-------------|------------|-------------------|--------|--------------|--------------------|---------|----------|----------|----------|
| LRN834545587    | 1084645587     | 22/10/2024    | 10846  | LEANDRO SOUZA ROSA      | CCM          | Sim                               | Sim                     | 10846       | LEANDRO SOUZA ROSA      | LRN8345   |                      | MK01       | OK         |                   | 10     | CCM          | Sim                | SG      | 0        | Log 2    | Log 3    |

**Descrição das colunas principais:**
- **conc**: Código de identificação.
- **conc2**: Outro código de identificação.
- **DATA**: Data do evento ou atividade.
- **MAT**: Matrícula do colaborador.
- **COLABORADOR**: Nome do colaborador.
- **PROCESSO**: Tipo de processo (ex.: CCM).
- **Foi para Campo Gerar Produção?**: Indicação se o colaborador foi ao campo gerar produção (Sim/Não).
- **Participou do DDS?**: Indicação se o colaborador participou do DDS (Sim/Não).
- **MAT ENC**: Matrícula do encarregado.
- **ENCARREGADO**: Nome do encarregado.
- **PLACA**: Placa do veículo.
- **PLACA INCORRETA?**: Indicação se a placa está incorreta (Sim/Não).
- **STATUS**: Status da atividade (ex.: OK).
- **JUSTIFICATIVA**: Justificativa do colaborador, se necessário.
- **MÊS**: Mês relacionado ao evento.
- **processo**: Tipo de processo repetido.
- **const base veic**: Não está claro, mas pode estar relacionado a um tipo de veículo ou código.
- **BASE**: Tipo de base ou categoria.
- **Log 1, Log 2, Log 3**: Logs do processo, ou registros de atividades adicionais.

---

### 2. `base_ps.xlsx` (Planilha de Colaboradores)

Esta planilha contém os registros dos colaboradores e seus horários de entrada e saída, bem como informações de processo e placa. Sua estrutura é composta pelas seguintes colunas:

| **DATA**      | **NOME**                                | **MATRICULA** | **HR IN**   | **HR FN**  | **AL IN**   | **AL FN**   | **PROCESSO** | **PLACA**   | **PROJETO** |
|---------------|-----------------------------------------|---------------|-------------|------------|-------------|-------------|--------------|-------------|-------------|
| 22/10/2024    | MARCO ANTONIO DO NASCIMENTO VALADARES   | 10849         | 7:30:00 AM  | 5:18:00 PM | 12:00:00 PM | 1:00:00 PM  | MANUTENÇÃO   | LMJ3E79     | 31124112103 |

**Descrição das colunas principais:**
- **DATA**: Data do evento ou jornada de trabalho.
- **NOME**: Nome completo do colaborador.
- **MATRICULA**: Matrícula do colaborador.
- **HR IN**: Hora de entrada no trabalho.
- **HR FN**: Hora de finalização do trabalho.
- **AL IN**: Hora de entrada para almoço.
- **AL FN**: Hora de finalização do almoço.
- **PROCESSO**: Tipo de processo que o colaborador está executando (ex.: Manutenção).
- **PLACA**: Placa do veículo associado ao colaborador.
- **PROJETO**: Código ou descrição do projeto relacionado.

---

## Funcionalidades do Sistema

O sistema foi desenvolvido para **facilitar o meu trabalho**, permitindo realizar as seguintes operações com os dados das planilhas:

- **Identificação de Colaboradores Sem Produção**: Comparar os colaboradores registrados na planilha `base_ps.xlsx` com os registrados na planilha `ps.xlsx` para identificar aqueles que não têm produção registrada.
- **Verificação de Dados Faltantes ou Incorretos**: Verificar inconsistências nos dados, como falta de produção ou erro nos dados de matrícula/placa.
- **Filtragem e Relatórios**: Gerar relatórios personalizados com base em filtros como data, matrícula, processo, entre outros.
- **Correção de Dados**: Permitir a correção de erros diretamente na interface, como marcar registros como "corrigidos".

## Como Utilizar

1. **Carregar as Planilhas**: Carregue as planilhas `ps.xlsx` e `base_ps.xlsx` no sistema para começar a análise.
2. **Aplicar Filtros**: Utilize os filtros disponíveis para selecionar dados específicos, como por data, colaborador ou tipo de processo.
3. **Gerar Relatórios**: Crie relatórios com base nas informações filtradas para análise ou exportação.
4. **Corrigir Dados**: Caso necessário, edite os registros diretamente nas planilhas ou no sistema para corrigir inconsistências.

## Requisitos

- Python 3.x
- Bibliotecas:
  - pandas
  - openpyxl

Para instalar as dependências, execute:

```bash
pip install -r requirements.txt
````

## Contribuições
Este sistema foi criado exclusivamente para facilitar o meu trabalho. No momento, não há necessidade de contribuições externas, mas se você tiver alguma sugestão ou feedback, sinta-se à vontade para entrar em contato.

## Licença
Este projeto está licenciado sob a Licença MIT.
