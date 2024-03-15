from openpyxl import load_workbook

# 1 - Lê pasta de trabalho e planilha
wb = load_workbook("data/pivot_table.xlsx")
sheet = wb["Relatorio"]

# 2 - Acessando um valor específico
#print(sheet["A3"].value)
#print(sheet["B3"].value)

# 3 - Iterando valores por meio de Loop
am_contador_total = 0 # meu
bt_contador_total = 0 # meu

for i in range(2, 6): # Professor
    ano = sheet["A%s" %i].value # Professor
    am = sheet["B%s" %i].value # Professor
    bt = sheet["C%s" %i].value # Professor
    print(f"Em {ano} o Aston Martin vendeu {am}") # Professor
    print(f"Em {ano} o Bentley vendeu {bt}") # Professor
    am_contador_total += am # meu
    bt_contador_total += bt # meu

    am_contador = 0 # meu
    bt_contador = 0 # meu 
    if (am > bt): # meu
        dfr = am - bt # meu
        print(f"Em {ano} Aston Martin vendeu {dfr} a mais que Bentley") # meu
        am_contador+=1 # meu
    else: # meu
        dfr = am - bt # meu
        print(f"Em {ano} Bentley vendeu {dfr} a mais que Aston Martin ") # meu
        bt_contador+=1 # meu
if am_contador > bt_contador: # meu
    print("Aston Martin vendeu mais que o Bentley") # meu
    print(f"O total foi {am_contador_total}") # meu
else: # meu
    print("Bentley vendeu mais que o Aston Martin") # meu