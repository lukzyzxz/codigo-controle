import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# -------------------- LOGIN --------------------
Usuarios = {
    "admin": "admin123",
    "proftec": "tecnico123",
    "professor": "prof123"
}

acesso_liberado = False
usuario_logado = None
while not acesso_liberado:
    usuario = input("\nDigite seu login: ").lower().strip()
    senha = input("Digite a senha: ").lower().strip()

    if usuario in Usuarios and Usuarios[usuario] == senha:
        print("\nLogin efetuado com sucesso")
        print("_________________________________\n")
        acesso_liberado = True
        usuario_logado = usuario
    else:
        print("\nLogin invalidado. Tente novamente")
        print("_________________________________")

# -------------------- CARREGAR DADOS --------------------
dados = None

if os.path.exists("alunos.csv"):
    dados = pd.read_csv("alunos.csv")
    print("✅ Dados carregados de alunos.csv")
elif os.path.exists("alunos.xlsx"):
    dados = pd.read_excel("alunos.xlsx")
    print("✅ Dados carregados de alunos.xlsx")
else:
    print("⚠ Nenhum arquivo de dados encontrado (alunos.csv ou alunos.xlsx)")
    exit()
os.system("cls")

# -------------------- FUNÇÕES --------------------
def exibir_informacoes():
    for idx, row in dados.iterrows():
        print("\n=== Dados do Computador ===")
        print(f"Número de Série do PC: {row['pc']}")
        print(f"Aluno: {row['nome']}")
        print(f"Horário de Entrada: {row['entrada']}")
        print(f"Horário de Saída: {row['saida']}")
        print("____________________________\n")

def gerar_relatorio():
    dados.to_csv("relatorio_alunos.csv", index=False)
    excel_path = "relatorio_alunos.xlsx"
    dados.to_excel(excel_path, index=False)
    wb = load_workbook(excel_path)
    ws = wb.active

    ws.insert_rows(1)
    ws["A1"] = "RELATÓRIO DE USO DOS COMPUTADORES"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)

    header_fill = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type="solid")
    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(excel_path)

    print("\n📂 Relatórios gerados com sucesso:")
    print(" - relatorio_alunos.csv")
    print(" - relatorio_alunos.xlsx (formatado)\n")

def menu_agendamento():
    arquivo_agendamentos = "agendamentos.csv"

    # Se não existir, cria com horários padrão
    if not os.path.exists(arquivo_agendamentos):
        horarios = [f"{h:02d}:00 - {h+1:02d}:00" for h in range(8, 21)]
        pcs = dados["pc"].unique()
        registros = []
        for pc in pcs:
            for h in horarios:
                registros.append([pc, h, "livre", "Disponível"])
        df_agend = pd.DataFrame(registros, columns=["pc", "horario", "professor", "status"])
        df_agend.to_csv(arquivo_agendamentos, index=False)

    # Carregar agendamentos
    df_agend = pd.read_csv(arquivo_agendamentos)

    print("\n=== AGENDAMENTOS ===")
    print("1 - Ver horários disponíveis")
    print("2 - Ver horários já agendados")
    print("3 - Agendar horário")
    print("4 - Voltar ao menu principal")

    escolha = input("Escolha uma opção: ")

    if escolha == "1":
        disponiveis = df_agend[df_agend["status"] == "Disponível"]
        if disponiveis.empty:
            print("\n⚠ Não há horários disponíveis")
        else:
            print("\nHorários disponíveis:")
            print(disponiveis[["pc", "horario"]])
    elif escolha == "2":
        agendados = df_agend[df_agend["status"] == "Agendado"]
        if agendados.empty:
            print("\n⚠ Nenhum horário agendado")
        else:
            print("\nHorários agendados:")
            print(agendados[["pc", "horario", "professor"]])
    elif escolha == "3":
        disponiveis = df_agend[df_agend["status"] == "Disponível"]
        if disponiveis.empty:
            print("\n⚠ Nenhum horário disponível para agendamento")
            return

        print("\nHorários disponíveis:")
        for i, row in disponiveis.iterrows():
            print(f"{i} - PC: {row['pc']} | Horário: {row['horario']}")

        escolha_idx = int(input("Digite o número do horário que deseja agendar: "))
        if escolha_idx in disponiveis.index:
            df_agend.loc[escolha_idx, "professor"] = usuario_logado
            df_agend.loc[escolha_idx, "status"] = "Agendado"
            df_agend.to_csv(arquivo_agendamentos, index=False)
            print("\n✅ Agendamento realizado com sucesso!")
        else:
            print("\n⚠ Opção inválida")
    elif escolha == "4":
        return
    else:
        print("\n⚠ Opção inválida")

# -------------------- MENU --------------------
if acesso_liberado:
    print("Escolha uma opção:")
    print("1 - Computadores")
    print("2 - Agendamento")
    print("3 - Relatório")

    opcao = input("Escolha o que deseja fazer: ")

    if opcao == "1":
        exibir_informacoes()
    elif opcao == "2":
        menu_agendamento()
    elif opcao == "3":
        gerar_relatorio()
    else:
        print("Opção inválida.")

Usuarios = {"login":{"super":123456}}