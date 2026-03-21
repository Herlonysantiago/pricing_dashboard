import customtkinter as ctk
import firebase_admin
from firebase_admin import credentials, db
import pandas as pd
from datetime import datetime
import tkinter.messagebox as messagebox
import os
import sys
# --- CONFIGURAÇÃO DOS CAMINHOS ---
def obter_caminho_recurso(nome_arquivo):
    """
    Se for a CHAVE (JSON), busca dentro do EXE (recurso fixo).
    Se for a PLANILHA (Excel), busca na PASTA do usuário (recurso editável).
    """
    if nome_arquivo.endswith(".json"):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, nome_arquivo)
    else:
        # Busca o Excel na mesma pasta onde o .exe está rodando
        return os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else "."), nome_arquivo)

# A Chave fica 'embutida' (segurança), a Planilha fica 'fora' (edição)
CAMINHO_CHAVE = obter_caminho_recurso("chave-firebase.json")
PLANILHA_INTERNA = obter_caminho_recurso("precos_internos.xlsx")
if not firebase_admin._apps:
    cred = credentials.Certificate(CAMINHO_CHAVE)
    firebase_admin.initialize_app(cred, {'databaseURL': 'https://pricing-ed61c-default-rtdb.firebaseio.com'})


class DashboardPricing(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PRICING - MONITOR DE COMBATE")
        self.geometry("1200x750")

        self.cor_azul = "#004a99"
        self.cor_laranja = "#f26522"
        self.cor_alerta = "#dc2626"

        self.dados_calculados = []
        self.mapa_produtos_internos = {}

        self.carregar_planilha_interna()

        # --- INTERFACE ---
        self.header = ctk.CTkFrame(self, height=80, fg_color=self.cor_azul, corner_radius=0)
        self.header.pack(fill="x")

        self.frame_filtros = ctk.CTkFrame(self.header, fg_color="transparent")
        self.frame_filtros.pack(side="left", padx=20)

        # Filtro Comprador
        lista_compradores = sorted(list(set([v['comprador'] for k, v in self.mapa_produtos_internos.items()])))
        compradores = ["TODOS COMPRADORES"] + lista_compradores
        self.combo_comprador = ctk.CTkComboBox(self.frame_filtros, values=compradores, width=200,
                                               command=lambda _: self.exibir_dados())
        self.combo_comprador.set("TODOS COMPRADORES")
        self.combo_comprador.grid(row=0, column=0, padx=5)

        # Busca Texto
        self.entry_busca = ctk.CTkEntry(self.frame_filtros, placeholder_text="Buscar produto...", width=280)
        self.entry_busca.grid(row=0, column=1, padx=5)
        self.entry_busca.bind("<KeyRelease>", lambda e: self.exibir_dados())

        # BOTÕES DE AÇÃO (Exportar e Atualizar)
        self.btn_export = ctk.CTkButton(self.header, text="EXPORTAR EXCEL", fg_color=self.cor_laranja,
                                        command=self.exportar_excel, width=140, font=("Arial", 12, "bold"))
        self.btn_export.pack(side="right", padx=20)

        self.btn_refresh = ctk.CTkButton(self.header, text="ATUALIZAR", fg_color="white", text_color=self.cor_azul,
                                         command=self.carregar_dados, width=100, font=("Arial", 12, "bold"))
        self.btn_refresh.pack(side="right", padx=10)

        self.main_frame = ctk.CTkScrollableFrame(self, fg_color="#f5f5f5")
        self.main_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.carregar_dados()

    def carregar_planilha_interna(self):
        try:
            df = pd.read_excel(PLANILHA_INTERNA, sheet_name='Pricing', skiprows=3)
            df.columns = [str(c).strip() for c in df.columns]

            for _, row in df.iterrows():
                cod_raw = str(row.get('Cod', '')).strip()
                if cod_raw and cod_raw != 'nan':
                    try:
                        codigo_limpo = str(int(float(cod_raw)))
                        emb = str(row.get('Embalagem', '')).strip()
                        if emb == 'nan': emb = ""

                        self.mapa_produtos_internos[codigo_limpo] = {
                            'preco': pd.to_numeric(row.get('VAREJO.1', 0), errors='coerce') or 0,
                            'comprador': str(row.get('Comprador', 'N/D')).strip(),
                            'embalagem': emb
                        }
                    except:
                        continue
        except Exception as e:
            print(f"Erro planilha: {e}")

    def carregar_dados(self):
        try:
            self.dados_calculados = []
            res = db.reference('produtos_lancados').get()
            if not res:
                self.exibir_dados()
                return

            for chave, p in res.items():
                cod_f = str(p.get('codigo', '')).strip()
                try:
                    cod_chave = str(int(float(cod_f)))
                except:
                    cod_chave = cod_f

                info_i = self.mapa_produtos_internos.get(cod_chave,
                                                         {'preco': 0, 'comprador': 'NÃO CADASTRADO', 'embalagem': ''})

                desc_final = f"{p.get('descricao', 'SEM DESCRIÇÃO')} - {info_i['embalagem']}" if info_i[
                    'embalagem'] else p.get('descricao', 'SEM DESCRIÇÃO')

                self.dados_calculados.append({
                    'CÓDIGO': cod_chave,
                    'DESCRIÇÃO': desc_final,
                    'MERCADO': p.get('mercado', 'N/D'),
                    'COMPRADOR': info_i['comprador'],
                    'NOSSO_PRECO': info_i['preco'],
                    'MERCADO_PRECO': float(p.get('preco', 0)),
                    'DIFERENCA_%': ((float(p.get('preco', 0)) - info_i['preco']) / info_i['preco'] * 100) if info_i[
                                                                                                                 'preco'] > 0 else 0
                })
            self.exibir_dados()
        except Exception as e:
            print(f"Erro Firebase: {e}")

    def exibir_dados(self):
        for widget in self.main_frame.winfo_children(): widget.destroy()

        headers = [("CÓD", 60), ("DESCRIÇÃO + EMBALAGEM", 400), ("MERCADO", 150), ("COMPRADOR", 180), ("NOSSO", 90),
                   ("MERC.", 90), ("DIF %", 80)]
        for i, (text, larg) in enumerate(headers):
            ctk.CTkLabel(self.main_frame, text=text, font=("Arial", 11, "bold"), text_color=self.cor_azul,
                         width=larg).grid(row=0, column=i, padx=5, pady=10)

        busca = self.entry_busca.get().upper().strip()
        f_comp = self.combo_comprador.get()

        row_idx = 1
        for d in self.dados_calculados:
            if (f_comp == "TODOS COMPRADORES" or d['COMPRADOR'] == f_comp) and (
                    not busca or busca in d['DESCRIÇÃO'].upper() or busca in d['CÓDIGO']):
                cor_dif = self.cor_alerta if d['NOSSO_PRECO'] > 0 and d['MERCADO_PRECO'] < d[
                    'NOSSO_PRECO'] else "#10b981"

                ctk.CTkLabel(self.main_frame, text=d['CÓDIGO']).grid(row=row_idx, column=0)
                ctk.CTkLabel(self.main_frame, text=d['DESCRIÇÃO'], font=("Arial", 10), wraplength=390,
                             justify="left").grid(row=row_idx, column=1, padx=5, pady=4)
                ctk.CTkLabel(self.main_frame, text=d['MERCADO'], font=("Arial", 10)).grid(row=row_idx, column=2)
                ctk.CTkLabel(self.main_frame, text=d['COMPRADOR'][:25], font=("Arial", 10)).grid(row=row_idx, column=3)
                ctk.CTkLabel(self.main_frame, text=f"R$ {d['NOSSO_PRECO']:.2f}").grid(row=row_idx, column=4)
                ctk.CTkLabel(self.main_frame, text=f"R$ {d['MERCADO_PRECO']:.2f}", font=("Arial", 11, "bold")).grid(
                    row=row_idx, column=5)
                ctk.CTkLabel(self.main_frame, text=f"{d['DIFERENCA_%']:+.2f}%", font=("Arial", 11, "bold"),
                             text_color=cor_dif).grid(row=row_idx, column=6)
                row_idx += 1

    def exportar_excel(self):
        """Exporta os dados filtrados ou todos para o Excel"""
        try:
            if not self.dados_calculados:
                messagebox.showwarning("Aviso", "Não há dados para exportar!")
                return

            # Filtra os dados de acordo com a busca e comprador atual antes de exportar
            busca = self.entry_busca.get().upper().strip()
            f_comp = self.combo_comprador.get()

            dados_para_exportar = [
                d for d in self.dados_calculados
                if (f_comp == "TODOS COMPRADORES" or d['COMPRADOR'] == f_comp) and
                   (not busca or busca in d['DESCRIÇÃO'].upper() or busca in d['CÓDIGO'])
            ]

            df_export = pd.DataFrame(dados_para_exportar)

            # Nome do arquivo com data e hora
            timestamp = datetime.now().strftime("%d_%m_%Y_%H%M")
            nome_arquivo = f"Relatorio_Combate_{timestamp}.xlsx"

            df_export.to_excel(nome_arquivo, index=False)
            messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso!\nSalvo como: {nome_arquivo}")

            # Tenta abrir a pasta onde o arquivo foi salvo (Linux/Mint)
            os.system(f'xdg-open .')

        except Exception as e:
            messagebox.showerror("Erro na Exportação", f"Não foi possível exportar: {e}")


if __name__ == "__main__":
    app = DashboardPricing()
    app.mainloop()