import tkinter as tk
from tkinter import ttk
import openpyxl

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

def load_data():
    path = "C:\\Users\\HP EliteBooK\\Desktop\\Stage\\Factures PR MFS 2024 4 check du 24 07 2024 (1).xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook['encours 2024']

    # Fetching all rows
    list_values = list(sheet.values)
    
    # Define the relevant columns you want to display (matching the treeview)
    relevant_columns = ("NO_CNC", "RA_SOCL", "CD_GRP_PROD", "NO_FACT_FOUR", "DT_PREM_FACT", "MT_LIG_FIN")

    # Set up the headings for treeview, ensuring only the relevant columns are included
    for col_name in relevant_columns:
        treeview.heading(col_name, text=col_name)
    
    # Insert rows into the treeview, filtering the values to match the relevant columns
    for value_tuple in list_values[1:]:
        # Create a tuple with only the relevant column data
        filtered_values = tuple(value_tuple[:len(relevant_columns)])
        treeview.insert('', tk.END, values=filtered_values)


def insert_row():
    no_cnc = no_cnc_entry.get()
    ra_soc = ra_soc_entry.get()
    cd_grp_prod = cd_grp_prod_entry.get()
    no_fact_four = no_fact_four_entry.get()
    dt_prem_fact = dt_prem_fact_entry.get()
    mt_lig_fin = float(mt_lig_fin_entry.get())

    print(no_cnc, ra_soc, cd_grp_prod, no_fact_four, dt_prem_fact, mt_lig_fin)

    # Insert row into Excel sheet
    path = "C:\\Users\\HP EliteBooK\\Desktop\\Stage\\Factures PR MFS 2024 4 check du 24 07 2024 (1).xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook['encours 2024']
    row_values = [no_cnc, ra_soc, cd_grp_prod, no_fact_four, dt_prem_fact, mt_lig_fin]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('', tk.END, values=row_values)
    
    # Clear the values
    no_cnc_entry.delete(0, "end")
    no_cnc_entry.insert(0, "NO_CNC")
    ra_soc_entry.delete(0, "end")
    ra_soc_entry.insert(0, "RA_SOCL")
    cd_grp_prod_entry.delete(0, "end")
    cd_grp_prod_entry.insert(0, "CD_GRP_PROD")
    no_fact_four_entry.delete(0, "end")
    no_fact_four_entry.insert(0, "NO_FACT_FOUR")
    dt_prem_fact_entry.delete(0, "end")
    dt_prem_fact_entry.insert(0, "DT_PREM_FACT")
    mt_lig_fin_entry.delete(0, "end")
    mt_lig_fin_entry.insert(0, "MT_LIG_FIN")

root = tk.Tk()

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

no_cnc_entry = ttk.Entry(widgets_frame)
no_cnc_entry.insert(0, "NO_CNC")
no_cnc_entry.bind("<FocusIn>", lambda e: no_cnc_entry.delete('0', 'end'))
no_cnc_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

ra_soc_entry = ttk.Entry(widgets_frame)
ra_soc_entry.insert(0, "RA_SOCL")
ra_soc_entry.bind("<FocusIn>", lambda e: ra_soc_entry.delete('0', 'end'))
ra_soc_entry.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

cd_grp_prod_entry = ttk.Entry(widgets_frame)
cd_grp_prod_entry.insert(0, "CD_GRP_PROD")
cd_grp_prod_entry.bind("<FocusIn>", lambda e: cd_grp_prod_entry.delete('0', 'end'))
cd_grp_prod_entry.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

no_fact_four_entry = ttk.Entry(widgets_frame)
no_fact_four_entry.insert(0, "NO_FACT_FOUR")
no_fact_four_entry.bind("<FocusIn>", lambda e: no_fact_four_entry.delete('0', 'end'))
no_fact_four_entry.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

dt_prem_fact_entry = ttk.Entry(widgets_frame)
dt_prem_fact_entry.insert(0, "DT_PREM_FACT")
dt_prem_fact_entry.bind("<FocusIn>", lambda e: dt_prem_fact_entry.delete('0', 'end'))
dt_prem_fact_entry.grid(row=4, column=0, padx=5, pady=5, sticky="ew")

mt_lig_fin_entry = ttk.Entry(widgets_frame)
mt_lig_fin_entry.insert(0, "MT_LIG_FIN")
mt_lig_fin_entry.bind("<FocusIn>", lambda e: mt_lig_fin_entry.delete('0', 'end'))
mt_lig_fin_entry.grid(row=5, column=0, padx=5, pady=5, sticky="ew")

button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=6, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=7, column=0, padx=(20, 10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(
    widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=8, column=0, padx=5, pady=10, sticky="nsew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("NO_CNC", "RA_SOCL", "CD_GRP_PROD", "NO_FACT_FOUR", "DT_PREM_FACT", "MT_LIG_FIN")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=13)
for col in cols:
    treeview.column(col, width=100)
treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()

root.mainloop()
