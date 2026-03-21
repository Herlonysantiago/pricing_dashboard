import customtkinter as ctk
import firebase_admin
from firebase_admin import credentials, db
import pandas as pd
from datetime import datetime
import tkinter.messagebox as messagebox

# 1. INICIALIZAÇÃO COM SUA CHAVE PRIVADA
cred = credentials.Certificate("chave-firebase.json")
firebase_admin.initialize_app(cred, {
    # Ajustado para o seu ID de projeto: pricing-ed61c
    'databaseURL': 'https://pricing-ed61c-default-rtdb.firebaseio.com'
})


class DashboardPricing(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("HS Pricing - Painel de Controle")
        self.geometry("1000x600")
        ctk.set_appearance_mode("light")

        # Cabeçalho
        self.header = ctk.CTkFrame(self, height=70, fg_color="#1e293b")
        self.header.pack(fill="x", padx=10, pady=10)

        self.lbl_title = ctk.CTkLabel(self.header, text="DASHBOARD PRICING", font=("Arial", 22, "bold"),
                                      text_color="white")
        self.lbl_title.pack(side="left", padx=20)

        self.btn_export = ctk.CTkButton(self.header, text="GERAR EXCEL", fg_color="#10b981", hover_color="#059669",
                                        command=self.exportar_excel)
        self.btn_export.pack(side="right", padx=10)

        self.btn_refresh = ctk.CTkButton(self.header, text="ATUALIZAR", command=self.carregar_dados)
        self.btn_refresh.pack(side="right", padx=10)

        # Container da Tabela
        self.main_frame = ctk.CTkScrollableFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.carregar_dados()

    def carregar_dados(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        # Títulos das Colunas
        cols = ["CÓDIGO", "DESCRIÇÃO", "PREÇO", "MERCADO", "AÇÃO"]
        for i, col in enumerate(cols):
            lbl = ctk.CTkLabel(self.main_frame, text=col, font=("Arial", 12, "bold"))
            lbl.grid(row=0, column=i, padx=20, pady=10, sticky="w")

        # Busca os produtos lançados (mesmo nó do Android)
        # Se o seu nó no Firebase for diferente de 'produtos_lancados', ajuste aqui:
        self.ref_lancados = db.reference('produtos_lancados').get()

        if self.ref_lancados:
            for idx, (chave, p) in enumerate(self.ref_lancados.items(), start=1):
                ctk.CTkLabel(self.main_frame, text=p.get('codigo', '-')).grid(row=idx, column=0, padx=20, pady=5,
                                                                              sticky="w")
                ctk.CTkLabel(self.main_frame, text=p.get('descricao', ''), wraplength=300).grid(row=idx, column=1,
                                                                                                padx=20, pady=5,
                                                                                                sticky="w")
                ctk.CTkLabel(self.main_frame, text=f"R$ {p.get('preco', '0')}".replace('.', ',')).grid(row=idx,
                                                                                                       column=2,
                                                                                                       padx=20, pady=5)
                ctk.CTkLabel(self.main_frame, text=p.get('mercado', '-')).grid(row=idx, column=3, padx=20, pady=5)

                # Botão para dar baixa na lista de "Faltam Lançar"
                btn = ctk.CTkButton(self.main_frame, text="BAIXAR", width=70,
                                    command=lambda c=p.get('codigo'): self.remover_pendente(c))
                btn.grid(row=idx, column=4, padx=20, pady=5)

    def remover_pendente(self, codigo):
        # Acessa a lista de pendentes e marca como pesquisado
        ref_pendentes = db.reference('produtos_pesquisa')
        snapshot = ref_pendentes.get()

        encontrou = False
        if snapshot:
            for chave, dados in snapshot.items():
                if str(dados.get('codigo')) == str(codigo):
                    ref_pendentes.child(chave).update({'pesquisado': True})
                    encontrou = True
                    break

        if encontrou:
            messagebox.showinfo("Sucesso", f"Item {codigo} removido da lista de faltas!")
            self.carregar_dados()
        else:
            messagebox.showwarning("Aviso", "Código não encontrado na lista de pendentes.")

    def exportar_excel(self):
        if not self.ref_lancados:
            return messagebox.showerror("Erro", "Não há dados para exportar.")

        try:
            df = pd.DataFrame(self.ref_lancados.values())
            nome_arq = f"Relatorio_Pricing_{datetime.now().strftime('%H%M%S')}.xlsx"
            df.to_excel(nome_arq, index=False)
            messagebox.showinfo("Sucesso", f"Excel gerado: {nome_arq}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar Excel: {e}")


if __name__ == "__main__":
    app = DashboardPricing()
    app.mainloop()