import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from Verificador_Produção_Colaboradores import top_level  # Importando a função top_level

# Classe para processar as planilhas
class ProcessadorPlanilha:
    def __init__(self, base_ps_path, ps_path):
        self.base_ps_path = base_ps_path
        self.ps_path = ps_path
        self.base_ps = None
        self.ps = None
        self.colaboradores_sem_producao = None
        self.preenchido_errado = None  # Para armazenar os registros preenchidos errado

    def processar_planilhas(self):
        # Carregar as planilhas
        self.base_ps = pd.read_excel(self.base_ps_path)
        self.ps = pd.read_excel(self.ps_path)

        # Certificar que as colunas de data são do tipo datetime
        self.base_ps['DATA'] = pd.to_datetime(self.base_ps['DATA'], errors='coerce')
        self.ps['DATA'] = pd.to_datetime(self.ps['DATA'], errors='coerce')

        # Criar a chave no DataFrame 'ps'
        self.ps['chave'] = self.ps['DATA'].astype(str) + self.ps['MAT'].astype(str)

        # Filtrar os registros da planilha 'ps' onde 'Foi para Campo Gerar Produção?' é 'Sim'
        ps_produtivos = self.ps[self.ps['Foi para Campo Gerar Produção?'] == 'Sim']

        # Criar a chave na base_ps
        self.base_ps['chave'] = self.base_ps['DATA'].astype(str) + self.base_ps['MATRICULA'].astype(str)

        # Filtrar para encontrar os registros que estão na base_ps, mas não na ps_produtivos
        self.colaboradores_sem_producao = self.base_ps[~self.base_ps['chave'].isin(ps_produtivos['chave'])]

        # Identificar registros em 'ps_produtivos' que não possuem correspondência na 'base_ps'
        self.preenchido_errado = ps_produtivos[~ps_produtivos['chave'].isin(self.base_ps['chave'])]

        # Remover registros duplicados com base na DATA e na MATRICULA, deixando apenas o primeiro
        self.colaboradores_sem_producao = self.colaboradores_sem_producao.drop_duplicates(subset=['DATA', 'MATRICULA'])

        # Combinar as duas listas de colaboradores: sem produção e preenchido errado
        colaboradores_final = pd.concat([self.colaboradores_sem_producao, self.preenchido_errado])

        # Criar a nova planilha
        colaboradores_final.to_excel('colaboradores_sem_producao.xlsx', index=False)

        return colaboradores_final


# Classe para a interface gráfica
class GUI:
    def __init__(self, root, processador):
        self.root = root
        self.root.title("Colaboradores Sem Produção - Compel")
        self.root.geometry("1000x700")  # Aumentando o tamanho da janela para 1000x700
        self.processador = processador
        self.treeview = None
        self.create_widgets()
        self.carregar_dados()

        # Bind para o evento de pressionar a tecla Enter
        self.root.bind("<Return>", self.marcar_corrigido)

    def create_widgets(self):
        # Filtros (Data e Matrícula)
        self.label_data = tk.Label(self.root, text="Data (dd/mm/yyyy):")
        self.label_data.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.entry_data = tk.Entry(self.root)
        self.entry_data.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        self.label_matricula = tk.Label(self.root, text="Matrícula:")
        self.label_matricula.grid(row=0, column=2, padx=10, pady=10, sticky="w")

        self.entry_matricula = tk.Entry(self.root)
        self.entry_matricula.grid(row=0, column=3, padx=10, pady=10, sticky="w")

        # Botão de filtro
        self.botao_filtro = tk.Button(self.root, text="Aplicar Filtro", command=self.aplicar_filtro)
        self.botao_filtro.grid(row=0, column=4, padx=10, pady=10)

        # Botão de reprocessamento das planilhas
        self.botao_reprocessar = tk.Button(self.root, text="Reprocessar Planilhas", command=self.reprocessar_planilhas)
        self.botao_reprocessar.grid(row=0, column=5, padx=10, pady=10)

        # Botão para abrir a janela do Verificador_Produção_Colaboradores
        self.botao_verificar = tk.Button(self.root, text="Abrir Verificador Produção", command=self.abrir_verificador_producao)
        self.botao_verificar.grid(row=0, column=6, padx=10, pady=10)

        # Criar o Treeview para exibir os dados
        columns = ['DATA', 'NOME', 'MATRICULA', 'PROCESSO', 'PLACA', 'Final de Semana', 'Corrigido']
        self.treeview = ttk.Treeview(self.root, columns=columns, show='headings')

        # Configurar as colunas
        self.treeview.heading('DATA', text="Data")
        self.treeview.heading('NOME', text="Nome")
        self.treeview.heading('MATRICULA', text="Matrícula")
        self.treeview.heading('PROCESSO', text="Processo")
        self.treeview.heading('PLACA', text="Placa")
        self.treeview.heading('Final de Semana', text="Final de Semana")
        self.treeview.heading('Corrigido', text="Corrigido")

        # Ajustar o tamanho das colunas para caber no conteúdo
        self.treeview.column('DATA', width=120, anchor='center')
        self.treeview.column('NOME', width=200, anchor='w')
        self.treeview.column('MATRICULA', width=100, anchor='center')
        self.treeview.column('PROCESSO', width=150, anchor='center')
        self.treeview.column('PLACA', width=100, anchor='center')
        self.treeview.column('Final de Semana', width=120, anchor='center')
        self.treeview.column('Corrigido', width=100, anchor='center')

        # Adicionar a barra de rolagem
        self.scrollbar_y = ttk.Scrollbar(self.root, orient="vertical", command=self.treeview.yview)
        self.treeview.configure(yscrollcommand=self.scrollbar_y.set)
        self.scrollbar_y.grid(row=1, column=7, sticky="ns", padx=10, pady=10)  # Ajuste na posição da barra de rolagem

        # Grid do Treeview
        self.treeview.grid(row=1, column=0, columnspan=7, padx=10, pady=10, sticky="nsew")

        # Botões de ação
        self.botao_corrigir = tk.Button(self.root, text="Marcar como Corrigido", command=self.marcar_corrigido)
        self.botao_corrigir.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        # Definindo as colunas da grid para ocupar a largura total
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_columnconfigure(3, weight=1)
        self.root.grid_columnconfigure(4, weight=1)
        self.root.grid_columnconfigure(5, weight=1)
        self.root.grid_columnconfigure(6, weight=1)

    def carregar_dados(self):
        # Carregar os dados processados
        self.dados = self.processador.colaboradores_sem_producao
        self.carregar_treeview(self.dados)

    def carregar_treeview(self, dados):
        # Limpar dados existentes no Treeview
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        # Adicionar os novos dados
        for _, row in dados.iterrows():
            # Formatar a data para 'dd/mm/yyyy'
            data_formatada = row['DATA'].strftime('%d/%m/%Y')
            # Verificar se é final de semana
            is_weekend = 'Sim' if row['DATA'].weekday() >= 5 else 'Não'  # Sábado (5) ou Domingo (6)
            # Determinar o estado da coluna 'Corrigido' (inicialmente 'Não')
            corrigido = 'Sim' if row.get('Corrigido', 'Não') == 'Sim' else 'Não'

            # Inserir dados no Treeview
            item_id = self.treeview.insert('', 'end', values=(data_formatada, row['NOME'], row['MATRICULA'], row['PROCESSO'], row['PLACA'], is_weekend, corrigido))

            # Alterar a cor da linha para verde se 'Corrigido' for 'Sim'
            if corrigido == 'Sim':
                self.treeview.item(item_id, tags=('corrigido',))

        # Aplicar a cor para as linhas com a tag 'corrigido'
        self.treeview.tag_configure('corrigido', background='lightgreen')

    def reprocessar_planilhas(self):
        # Reprocessar as planilhas
        self.processador.processar_planilhas()

        # Atualizar a interface com os novos dados
        self.carregar_dados()

        messagebox.showinfo("Sucesso", "Planilhas reprocessadas com sucesso!")

    def marcar_corrigido(self, event=None):
        # Marcar os registros selecionados como corrigidos
        selected_items = self.treeview.selection()
        for item in selected_items:
            values = self.treeview.item(item)['values']
            # Atualizar a coluna "Corrigido" para 'Sim' para os itens selecionados
            self.treeview.item(item, values=(values[0], values[1], values[2], values[3], values[4], values[5], 'Sim'))

            # Alterar a cor da linha para verde
            self.treeview.item(item, tags=('corrigido',))

    def aplicar_filtro(self):
        data_input = self.entry_data.get()
        matricula_input = self.entry_matricula.get()

        # Filtrar os dados conforme os valores de entrada
        dados_filtrados = self.dados

        if data_input:
            try:
                data_formatada = pd.to_datetime(data_input, format='%d/%m/%Y')
                dados_filtrados = dados_filtrados[dados_filtrados['DATA'] == data_formatada]
            except ValueError:
                messagebox.showerror("Erro", "Formato de data inválido. Use dd/mm/yyyy.")

        if matricula_input:
            dados_filtrados = dados_filtrados[dados_filtrados['MATRICULA'].astype(str) == matricula_input]

        self.carregar_treeview(dados_filtrados)

    def abrir_verificador_producao(self):
        # Chama a função top_level de Verificador_Produção_Colaboradores
        top_level(self.root)


# Função principal
def main():
    # Caminhos das planilhas de entrada
    base_ps_path = 'base_ps.xlsx'
    ps_path = 'ps.xlsx'

    # Criar objeto de processamento de planilhas
    processador = ProcessadorPlanilha(base_ps_path, ps_path)

    # Processar as planilhas assim que o programa iniciar
    processador.processar_planilhas()

    # Criar a interface gráfica
    root = tk.Tk()
    gui = GUI(root, processador)
    root.mainloop()

# Executar a função principal
if __name__ == "__main__":
    main()
