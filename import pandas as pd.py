import pandas as pd
import os
import pwinput
import time
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
    senha = pwinput.pwinput(prompt="Digite a senha: ", mask="*").lower().strip()

    # limpa a tela logo após digitar a senha
    os.system("cls" if os.name == "nt" else "clear")

    if usuario in Usuarios and Usuarios[usuario] == senha:
        print("\n✅ Login efetuado com sucesso")
        print("_________________________________\n")
        acesso_liberado = True
        usuario_logado = usuario
        time.sleep(1)
    else:
        print("\n❌ Login inválido. Tente novamente")
        print("_________________________________")
        time.sleep(1)
        os.system("cls" if os.name == "nt" else "clear")


# -------------------- MENU COMPUTADORES --------------------
def menu_computadores():
    print("\n=== MENU COMPUTADORES ===")
    print("1 - Registrar novo aluno")
    print("2 - Consultar alunos cadastrados")
    print("3 - Editar aluno")
    print("4 - Excluir aluno")
    print("5 - Voltar ao menu principal")

    escolha = input("Escolha uma opção: ")

    # -------------------- REGISTRAR --------------------
    if escolha == "1":
        pc = input("Digite o número de série do PC: ").strip()
        nome = input("Digite o nome do aluno: ").strip()
        data = input("Digite a data (DD/MM/AAAA): ").strip()  # NOVO CAMPO
        entrada = input("Digite o horário de entrada (HH:MM): ").strip()
        saida = input("Digite o horário de saída (HH:MM): ").strip()

        print("\n💾 Salvando registro...")
        time.sleep(1.2)

        novo_registro = pd.DataFrame([{
            "pc": pc,
            "nome": nome,
            "data": data,
            "entrada": entrada,
            "saida": saida
        }])

        arquivo_csv = "alunos.csv"
        arquivo_xlsx = "alunos.xlsx"

        if os.path.exists(arquivo_csv):
            antigo = pd.read_csv(arquivo_csv)
            atualizado = pd.concat([antigo, novo_registro], ignore_index=True)
            atualizado.to_csv(arquivo_csv, index=False)
            atualizado.to_excel(arquivo_xlsx, index=False)
        else:
            novo_registro.to_csv(arquivo_csv, index=False)
            novo_registro.to_excel(arquivo_xlsx, index=False)

        print("\n✅ Registro salvo com sucesso!")

    # -------------------- CONSULTAR --------------------
    elif escolha == "2":
        print("\n📂 Carregando dados dos alunos...")
        time.sleep(1)

        if os.path.exists("alunos.csv"):
            dados = pd.read_csv("alunos.csv")
            if dados.empty:
                print("\n⚠ Nenhum aluno cadastrado ainda.")
            else:
                for idx, row in dados.iterrows():
                    print("\n=== Registro", idx, "===")
                    print(f"Número de Série do PC: {row['pc']}")
                    print(f"Aluno: {row['nome']}")
                    print(f"Data: {row['data']}")
                    print(f"Horário de Entrada: {row['entrada']}")
                    print(f"Horário de Saída: {row['saida']}")
                    print("____________________________")
        else:
            print("\n⚠ Nenhum arquivo de alunos encontrado.")

    # -------------------- EDITAR --------------------
    elif escolha == "3":
        if not os.path.exists("alunos.csv"):
            print("\n⚠ Nenhum arquivo encontrado para edição.")
            return

        dados = pd.read_csv("alunos.csv")
        if dados.empty:
            print("\n⚠ Nenhum aluno cadastrado para editar.")
            return

        print("\nAlunos cadastrados:")
        print(dados[["pc", "nome", "data"]])

        try:
            idx = int(input("Digite o índice do aluno que deseja editar: "))
            if idx not in dados.index:
                print("\n⚠ Índice inválido.")
                return
        except ValueError:
            print("\n⚠ Entrada inválida.")
            return

        print("\nDeixe em branco para não alterar.")
        novo_pc = input(f"PC atual ({dados.loc[idx,'pc']}): ").strip()
        novo_nome = input(f"Nome atual ({dados.loc[idx,'nome']}): ").strip()
        nova_data = input(f"Data atual ({dados.loc[idx,'data']}): ").strip()
        nova_entrada = input(f"Entrada atual ({dados.loc[idx,'entrada']}): ").strip()
        nova_saida = input(f"Saída atual ({dados.loc[idx,'saida']}): ").strip()

        print("\n🔄 Atualizando registro...")
        time.sleep(1.3)

        if novo_pc: dados.loc[idx,"pc"] = novo_pc
        if novo_nome: dados.loc[idx,"nome"] = novo_nome
        if nova_data: dados.loc[idx,"data"] = nova_data
        if nova_entrada: dados.loc[idx,"entrada"] = nova_entrada
        if nova_saida: dados.loc[idx,"saida"] = nova_saida

        dados.to_csv("alunos.csv", index=False)
        dados.to_excel("alunos.xlsx", index=False)

        print("\n✅ Registro atualizado com sucesso!")

    # -------------------- EXCLUIR --------------------
    elif escolha == "4":
        if not os.path.exists("alunos.csv"):
            print("\n⚠ Nenhum arquivo encontrado para exclusão.")
            return

        dados = pd.read_csv("alunos.csv")
        if dados.empty:
            print("\n⚠ Nenhum aluno cadastrado para excluir.")
            return

        print("\nAlunos cadastrados:")
        print(dados[["pc", "nome", "data"]])

        try:
            idx = int(input("Digite o índice do aluno que deseja excluir: "))
            if idx not in dados.index:
                print("\n⚠ Índice inválido.")
                return
        except ValueError:
            print("\n⚠ Entrada inválida.")
            return

        confirmacao = input(f"Tem certeza que deseja excluir o registro de {dados.loc[idx,'nome']} no dia {dados.loc[idx,'data']}? (s/n): ").lower()
        if confirmacao == "s":
            print("\n🗑 Apagando registro...")
            time.sleep(1)

            dados = dados.drop(idx).reset_index(drop=True)
            dados.to_csv("alunos.csv", index=False)
            dados.to_excel("alunos.xlsx", index=False)

            print("\n✅ Registro excluído com sucesso!")
        else:
            print("\n⚠ Exclusão cancelada.")

    # -------------------- VOLTAR --------------------
    elif escolha == "5":
        return
    else:
        print("\n⚠ Opção inválida")


# -------------------- FUNÇÃO RELATÓRIO --------------------
def gerar_relatorio():
    print("\n=== RELATÓRIO DE AULA ===")
    professor = input("Digite seu nome (professor): ").strip()
    descricao = input("Digite o relatório da aula: ").strip()

    print("\n💾 Salvando relatório...")
    time.sleep(1.3)

    novo_relatorio = pd.DataFrame([{
        "professor": professor,
        "relatorio": descricao
    }])

    arquivo_csv = "relatorios.csv"
    arquivo_xlsx = "relatorios.xlsx"

    if os.path.exists(arquivo_csv):
        antigo = pd.read_csv(arquivo_csv)
        atualizado = pd.concat([antigo, novo_relatorio], ignore_index=True)
        atualizado.to_csv(arquivo_csv, index=False)
        atualizado.to_excel(arquivo_xlsx, index=False)
    else:
        novo_relatorio.to_csv(arquivo_csv, index=False)
        novo_relatorio.to_excel(arquivo_xlsx, index=False)

    print("\n✅ Relatório salvo com sucesso!")


# -------------------- AGENDAMENTOS --------------------
def menu_agendamento():
    arquivo_agendamentos = "agendamentos.csv"

    if not os.path.exists(arquivo_agendamentos):
        horarios = [f"{h:02d}:00 - {h+1:02d}:00" for h in range(8, 21)]
        pcs = ["PC01", "PC02", "PC03"]
        registros = []
        for pc in pcs:
            for h in horarios:
                registros.append([pc, h, "livre", "Disponível"])
        df_agend = pd.DataFrame(registros, columns=["pc", "horario", "professor", "status"])
        df_agend.to_csv(arquivo_agendamentos, index=False)

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
            print("\n🔄 Reservando horário...")
            time.sleep(1.5)

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


# -------------------- LIMPAR DADOS (ADMIN) --------------------
def limpar_dados():
    if usuario_logado != "admin":
        print("\n⚠ Apenas o ADMIN pode acessar esta opção.")
        return

    print("\n=== MENU DE LIMPEZA DE DADOS ===")
    print("1 - Limpar relatórios")
    print("2 - Limpar agendamentos")
    print("3 - Limpar alunos")
    print("4 - Limpar tudo")
    print("5 - Voltar")

    escolha = input("Escolha uma opção: ")

    if escolha == "1":
        print("\n🧹 Limpando relatórios...")
        time.sleep(1.5)
        for arq in ["relatorios.csv", "relatorios.xlsx"]:
            if os.path.exists(arq): os.remove(arq)
        print("\n✅ Relatórios apagados com sucesso!")

    elif escolha == "2":
        print("\n🧹 Limpando agendamentos...")
        time.sleep(1.5)
        if os.path.exists("agendamentos.csv"):
            os.remove("agendamentos.csv")
        print("\n✅ Agendamentos apagados com sucesso!")

    elif escolha == "3":
        print("\n🧹 Limpando alunos...")
        time.sleep(1.5)
        for arq in ["alunos.csv", "alunos.xlsx"]:
            if os.path.exists(arq): os.remove(arq)
        print("\n✅ Alunos apagados com sucesso!")

    elif escolha == "4":
        print("\n🧹 Limpando todos os dados do sistema...")
        time.sleep(2)
        for arq in ["relatorios.csv", "relatorios.xlsx", "agendamentos.csv", "alunos.csv", "alunos.xlsx"]:
            if os.path.exists(arq): os.remove(arq)
        print("\n✅ Todos os dados foram apagados com sucesso!")

    elif escolha == "5":
        return
    else:
        print("\n⚠ Opção inválida.")


# -------------------- MENU PRINCIPAL --------------------
if acesso_liberado:
    while True:
        print("\n=== MENU PRINCIPAL ===")
        print("1 - Computadores")
        print("2 - Agendamento")
        print("3 - Relatório")
        print("4 - Sair")
        if usuario_logado == "admin":
            print("5 - Limpar dados")

        opcao = input("Escolha o que deseja fazer: ")

        if opcao == "1":
            menu_computadores()
        elif opcao == "2":
            menu_agendamento()
        elif opcao == "3":
            gerar_relatorio()
        elif opcao == "4":
            print("\n👋 Saindo do sistema...")
            time.sleep(1)
            break
        elif opcao == "5" and usuario_logado == "admin":
            limpar_dados()
        else:
            print("\n⚠ Opção inválida.")
