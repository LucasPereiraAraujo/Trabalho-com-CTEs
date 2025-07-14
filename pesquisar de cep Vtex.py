import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def carregar_planilha():
    caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if caminho:
        try:
            df = pd.read_excel(caminho)
            df['ZipCodeStart'] = df['ZipCodeStart'].astype(int)
            df['ZipCodeEnd'] = df['ZipCodeEnd'].astype(int)
            df['WeightStart'] = df['WeightStart'].astype(float)
            df['WeightEnd'] = df['WeightEnd'].astype(float)
            app.df = df
            app.planilha_nome = os.path.basename(caminho)
            planilha_label.config(text=f"Planilha: {app.planilha_nome}")
            messagebox.showinfo("Sucesso", "Planilha carregada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar planilha: {e}")

def calcular_frete():
    if not hasattr(app, 'df'):
        messagebox.showerror("Erro", "Carregue a planilha primeiro.")
        return

    try:
        cep = int(cep_entry.get())
        peso = float(peso_entry.get())
        acrescimo = float(acrescimo_entry.get()) if acrescimo_entry.get() else 0
        valor_encomenda = float(valor_entry.get()) if valor_entry.get() else 0

        peso_final = peso * (1 + acrescimo / 100)

        # Filtra a faixa correspondente
        df_filtrado = app.df[
            (app.df['ZipCodeStart'] <= cep) & 
            (app.df['ZipCodeEnd'] >= cep) &
            (app.df['WeightStart'] <= peso_final) & 
            (app.df['WeightEnd'] >= peso_final)
        ]

        if df_filtrado.empty:
            resultado_var.set("Nenhuma faixa encontrada.")
            return

        linha = df_filtrado.iloc[0]
        valor_frete = linha['AbsoluteMoneyCost']
        prazo_entrega = int(str(linha['TimeCost']).split('.')[0])

        if valor_encomenda > 0:
            seguro = (linha['PricePercent'] / 100) * valor_encomenda
        else:
            seguro = 0

        total = valor_frete + seguro
        resultado_var.set(
            f"Frete: R${valor_frete:.2f} | Seguro: R${seguro:.2f} | "
            f"Total: R${total:.2f} | Prazo: {prazo_entrega} dias"
        )

    except Exception as e:
        resultado_var.set(f"Erro: {e}")

# Interface
app = tk.Tk()
app.title("Calculadora de Frete por CEP e Peso")

tk.Button(app, text="Carregar Planilha", command=carregar_planilha).grid(row=0, column=0, columnspan=2, pady=10)

planilha_label = tk.Label(app, text="Planilha: (nenhuma carregada)", fg="blue")
planilha_label.grid(row=1, column=0, columnspan=2)

tk.Label(app, text="CEP:").grid(row=2, column=0)
cep_entry = tk.Entry(app)
cep_entry.grid(row=2, column=1)

tk.Label(app, text="Peso (g):").grid(row=3, column=0)
peso_entry = tk.Entry(app)
peso_entry.grid(row=3, column=1)

tk.Label(app, text="Acréscimo de Peso (%):").grid(row=4, column=0)
acrescimo_entry = tk.Entry(app)
acrescimo_entry.grid(row=4, column=1)

tk.Label(app, text="Valor da Encomenda (R$):").grid(row=5, column=0)
valor_entry = tk.Entry(app)
valor_entry.grid(row=5, column=1)

tk.Button(app, text="Calcular Frete", command=calcular_frete).grid(row=6, column=0, columnspan=2, pady=10)

resultado_var = tk.StringVar()
tk.Label(app, textvariable=resultado_var, font=('Arial', 12, 'bold'), wraplength=400, justify="center").grid(row=7, column=0, columnspan=2, pady=10)

app.mainloop()
