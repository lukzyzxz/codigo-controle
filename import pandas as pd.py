import pandas as pd
import os
import pwinput
import time
from datetime import datetime, timedelta
from tabulate import tabulate
from pathlib import Path
from colorama import Fore, Style, init
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

init(autoreset=True)

ARQ_ALUNOS = Path("alunos.csv")
ARQ_ALUNOS_XLSX = Path("alunos.xlsx")
ARQ_REL = Path("relatorios.csv")
ARQ_REL_XLSX = Path("relatorios.xlsx")
ARQ_AG = Path("agendamentos.csv")
USERS = {
    "admin": {"senha": "admin123", "nome": "Administrador"},
    "proftec": {"senha": "tecnico123", "nome": "Prof. Técnico"},
    "professor": {"senha": "prof123", "nome": "Professor"}
}

def limpar_tela():
    os.system("cls" if os.name == "nt" else "clear")

def msg(text, tipo="info"):
    cores = {"info": Fore.CYAN, "ok": Fore.GREEN, "warn": Fore.YELLOW, "err": Fore.RED}
    print(cores.get(tipo, Fore.CYAN) + text + Style.RESET_ALL)

def pedir_validado(prompt, func):
    while True:
        v = input(prompt).strip()
        r = func(v)
        if r is not None:
            return r

def validar_nome(nome: str):
    if nome.replace(" ", "").isalpha():
        return nome.title().strip()
    return None

def validar_numero(texto: str):
    if texto.isdigit():
        return texto
    return None

def validar_data(data_str: str):
    try:
        d = datetime.strptime(data_str, "%d/%m/%Y")
        return d.strftime("%d/%m/%Y")
    except Exception:
        return None

def validar_hora(hora_str: str):
    try:
        h = datetime.strptime(hora_str, "%H:%M")
        return h.strftime("%H:%M")
    except Exception:
        return None

def confirmar_sn(mensagem: str):
    while True:
        r = input(mensagem + " (s/n): ").lower().strip()
        if r in ("s", "n"):
            return r

def salvar_csv_xlsx(df: pd.DataFrame, csv_path: Path, xlsx_path: Path):
    df.to_csv(csv_path, index=False)
    try:
        df.to_excel(xlsx_path, index=False)
        try:
            wb = load_workbook(xlsx_path)
            ws = wb.active
            header_font = Font(bold=True)
            for cell in next(ws.iter_rows(min_row=1, max_row=1)):
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            wb.save(xlsx_path)
        except Exception:
            pass
    except Exception:
        df.to_csv(csv_path, index=False)

def calcular_duracao(data_str: str, entrada_str: str, saida_str: str) -> str:
    try:
        data = datetime.strptime(data_str, "%d/%m/%Y")
        entrada = datetime.strptime(entrada_str, "%H:%M").time()
        saida = datetime.strptime(saida_str, "%H:%M").time()
        dt_entrada = datetime.combine(data.date(), entrada)
        dt_saida = datetime.combine(data.date(), saida)
        if dt_saida < dt_entrada:
            dt_saida += timedelta(days=1)
        delta = dt_saida - dt_entrada
        horas = int(delta.total_seconds() // 3600)
        minutos = int((delta.total_seconds() % 3600) // 60)
        return f"{horas:02d}:{minutos:02d}"
    except Exception:
        return ""

def carregar_dataframe(path: Path, cols=None):
    if not path.exists():
        if cols:
            return pd.DataFrame(columns=cols)
        return pd.DataFrame()
    try:
        return pd.read_csv(path)
    except Exception:
        return pd.DataFrame()

def menu_computadores(usuario_logado: str):
    while True:
        print("\n=== MENU COMPUTADORES ===")
        print("1 - Registrar novo aluno")
        print("2 - Consultar alunos cadastrados")
        print("3 - Editar aluno")
        print("4 - Excluir aluno")
        print("5 - Voltar ao menu principal")
        escolha = input("Escolha uma opção: ").strip()

        if escolha == "1":
            numero_pc = pedir_validado("Digite o número do PC (ex: 01, 02...): ", validar_numero)
            pc = f"PC{numero_pc.zfill(2)}"
            nome = pedir_validado("Digite o nome do aluno: ", validar_nome)

            print("\nDeseja usar a data e hora atuais para o registro?")
            print("1 - Sim, usar data e hora atuais")
            print("2 - Não, quero inserir manualmente")
            opc = ""
            while opc not in ("1", "2"):
                opc = input("Escolha uma opção (1/2): ").strip()

            if opc == "1":
                agora = datetime.now()
                data_automatica = agora.strftime("%d/%m/%Y")
                hora_automatica = agora.strftime("%H:%M")
                print(f"\n📅 Data atual: {data_automatica}")
                print(f"🕒 Horário atual: {hora_automatica}")
                if confirmar_sn("Deseja confirmar essa data e hora?") == "s":
                    data = data_automatica
                    entrada = hora_automatica
                    msg("\n✅ Data e hora registradas automaticamente.", "ok")
                else:
                    data = None
                    entrada = None
            else:
                data = None
                entrada = None

            if not data:
                data = pedir_validado("Digite a data (DD/MM/AAAA): ", validar_data)
            if not entrada:
                entrada = pedir_validado("Digite o horário de entrada (HH:MM): ", validar_hora)
            saida = pedir_validado("Digite o horário de saída (HH:MM): ", validar_hora)

            duracao = calcular_duracao(data, entrada, saida)

            novo_registro = pd.DataFrame([{
                "pc": pc,
                "nome": nome,
                "data": data,
                "entrada": entrada,
                "saida": saida,
                "duracao": duracao
            }])

            df = carregar_dataframe(ARQ_ALUNOS, cols=["pc", "nome", "data", "entrada", "saida", "duracao"])
            if not df.empty:
                df = pd.concat([df, novo_registro], ignore_index=True)
            else:
                df = novo_registro

            salvar_csv_xlsx(df, ARQ_ALUNOS, ARQ_ALUNOS_XLSX)
            msg("\n✅ Registro salvo com sucesso!", "ok")

        elif escolha == "2":
            msg("\n📂 Carregando dados dos alunos...", "info")
            time.sleep(0.6)
            dados = carregar_dataframe(ARQ_ALUNOS)
            if dados.empty:
                msg("\n⚠ Nenhum aluno cadastrado ainda.", "warn")
            else:
                print("\n=== Alunos Cadastrados ===")
                print(tabulate(dados, headers="keys", tablefmt="grid", showindex=True))

        elif escolha == "3":
            if not ARQ_ALUNOS.exists():
                msg("\n⚠ Nenhum arquivo encontrado para edição.", "warn")
                continue
            dados = carregar_dataframe(ARQ_ALUNOS)
            if dados.empty:
                msg("\n⚠ Nenhum aluno cadastrado para editar.", "warn")
                continue
            print("\nAlunos cadastrados:")
            print(tabulate(dados, headers="keys", tablefmt="grid", showindex=True))
            try:
                idx = int(input("Digite o índice do aluno que deseja editar: ").strip())
            except Exception:
                msg("\n⚠ Entrada inválida.", "warn")
                continue
            if idx not in dados.index:
                msg("\n⚠ Índice inválido.", "warn")
                continue
            print("\nDeixe em branco para não alterar.")
            novo_pc = input(f"PC atual ({dados.loc[idx,'pc']}): ").strip()
            if novo_pc:
                num = validar_numero(novo_pc.replace("PC", "").replace("pc", ""))
                if num:
                    novo_pc = f"PC{num.zfill(2)}"
                else:
                    novo_pc = dados.loc[idx, "pc"]
            else:
                novo_pc = dados.loc[idx, "pc"]

            novo_nome = input(f"Nome atual ({dados.loc[idx,'nome']}): ").strip()
            if novo_nome:
                valid = validar_nome(novo_nome)
                novo_nome = valid if valid else dados.loc[idx, "nome"]
            else:
                novo_nome = dados.loc[idx, "nome"]

            nova_data = input(f"Data atual ({dados.loc[idx,'data']}): ").strip()
            if nova_data:
                nova_data = validar_data(nova_data) or dados.loc[idx, "data"]
            else:
                nova_data = dados.loc[idx, "data"]

            nova_entrada = input(f"Entrada atual ({dados.loc[idx,'entrada']}): ").strip()
            if nova_entrada:
                nova_entrada = validar_hora(nova_entrada) or dados.loc[idx, "entrada"]
            else:
                nova_entrada = dados.loc[idx, "entrada"]

            nova_saida = input(f"Saída atual ({dados.loc[idx,'saida']}): ").strip()
            if nova_saida:
                nova_saida = validar_hora(nova_saida) or dados.loc[idx, "saida"]
            else:
                nova_saida = dados.loc[idx, "saida"]

            duracao = calcular_duracao(nova_data, nova_entrada, nova_saida)

            dados.loc[idx, "pc"] = novo_pc
            dados.loc[idx, "nome"] = novo_nome
            dados.loc[idx, "data"] = nova_data
            dados.loc[idx, "entrada"] = nova_entrada
            dados.loc[idx, "saida"] = nova_saida
            dados.loc[idx, "duracao"] = duracao

            salvar_csv_xlsx(dados, ARQ_ALUNOS, ARQ_ALUNOS_XLSX)
            msg("\n✅ Registro atualizado com sucesso!", "ok")

        elif escolha == "4":
            if not ARQ_ALUNOS.exists():
                msg("\n⚠ Nenhum arquivo encontrado para exclusão.", "warn")
                continue
            dados = carregar_dataframe(ARQ_ALUNOS)
            if dados.empty:
                msg("\n⚠ Nenhum aluno cadastrado para excluir.", "warn")
                continue
            print("\nAlunos cadastrados:")
            print(tabulate(dados, headers="keys", tablefmt="grid", showindex=True))
            try:
                idx = int(input("Digite o índice do aluno que deseja excluir: ").strip())
            except Exception:
                msg("\n⚠ Entrada inválida.", "warn")
                continue
            if idx not in dados.index:
                msg("\n⚠ Índice inválido.", "warn")
                continue
            if confirmar_sn(f"Tem certeza que deseja excluir o registro de {dados.loc[idx,'nome']} no dia {dados.loc[idx,'data']}?") == "s":
                msg("\n🗑 Apagando registro...", "info")
                time.sleep(0.6)
                dados = dados.drop(idx).reset_index(drop=True)
                salvar_csv_xlsx(dados, ARQ_ALUNOS, ARQ_ALUNOS_XLSX)
                msg("\n✅ Registro excluído com sucesso!", "ok")
            else:
                msg("\n⚠ Exclusão cancelada.", "warn")

        elif escolha == "5":
            return
        else:
            msg("\n⚠ Opção inválida.", "warn")

def gerar_relatorio(usuario_logado: str):
    print("\n=== RELATÓRIO DE AULA ===")
    professor = pedir_validado("Digite seu nome (professor): ", validar_nome)
    descricao = input("Digite o relatório da aula: ").strip()
    msg("\n💾 Salvando relatório...", "info")
    time.sleep(0.6)
    novo_rel = pd.DataFrame([{"professor": professor, "relatorio": descricao, "usuario": usuario_logado}])
    df = carregar_dataframe(ARQ_REL, cols=["professor", "relatorio", "usuario"])
    if not df.empty:
        df = pd.concat([df, novo_rel], ignore_index=True)
    else:
        df = novo_rel
    salvar_csv_xlsx(df, ARQ_REL, ARQ_REL_XLSX)
    msg("\n✅ Relatório salvo com sucesso!", "ok")

def menu_agendamento(usuario_logado: str):
    if not ARQ_AG.exists():
        horarios = [f"{h:02d}:00 - {h+1:02d}:00" for h in range(8, 21)]
        pcs = ["PC01", "PC02", "PC03"]
        registros = []
        for pc in pcs:
            for h in horarios:
                registros.append([pc, h, "livre", "Disponível"])
        df_ag = pd.DataFrame(registros, columns=["pc", "horario", "professor", "status"])
        salvar_csv_xlsx(df_ag, ARQ_AG, Path("agendamentos.xlsx"))
    df_agend = carregar_dataframe(ARQ_AG, cols=["pc", "horario", "professor", "status"])
    print("\n=== AGENDAMENTOS ===")
    print("1 - Ver horários disponíveis")
    print("2 - Ver horários já agendados")
    print("3 - Agendar horário")
    print("4 - Voltar ao menu principal")
    escolha = input("Escolha uma opção: ").strip()

    if escolha == "1":
        disponiveis = df_agend[df_agend["status"] == "Disponível"]
        if disponiveis.empty:
            msg("\n⚠ Não há horários disponíveis", "warn")
        else:
            print("\nHorários disponíveis:")
            print(tabulate(disponiveis[["pc", "horario"]], headers="keys", tablefmt="grid", showindex=True))

    elif escolha == "2":
        agendados = df_agend[df_agend["status"] == "Agendado"]
        if agendados.empty:
            msg("\n⚠ Nenhum horário agendado", "warn")
        else:
            print("\nHorários agendados:")
            print(tabulate(agendados[["pc", "horario", "professor"]], headers="keys", tablefmt="grid", showindex=True))

    elif escolha == "3":
        disponiveis = df_agend[df_agend["status"] == "Disponível"]
        if disponiveis.empty:
            msg("\n⚠ Nenhum horário disponível para agendamento", "warn")
            return
        print("\nHorários disponíveis:")
        print(tabulate(disponiveis[["pc", "horario"]], headers="keys", tablefmt="grid", showindex=True))
        try:
            escolha_idx = int(input("Digite o número do horário que deseja agendar: ").strip())
        except Exception:
            msg("\n⚠ Entrada inválida.", "warn")
            return
        if escolha_idx in disponiveis.index:
            horario_escolhido = df_agend.loc[escolha_idx, "horario"]
            conflito = df_agend[
                (df_agend["professor"] == USERS[usuario_logado]["nome"]) &
                (df_agend["horario"] == horario_escolhido) &
                (df_agend["status"] == "Agendado")
            ]
            if not conflito.empty:
                msg("\n⚠ Você já tem um agendamento nesse mesmo horário.", "warn")
                return
            msg("\n🔄 Reservando horário...", "info")
            time.sleep(0.6)
            df_agend.loc[escolha_idx, "professor"] = USERS[usuario_logado]["nome"]
            df_agend.loc[escolha_idx, "status"] = "Agendado"
            salvar_csv_xlsx(df_agend, ARQ_AG, Path("agendamentos.xlsx"))
            msg("\n✅ Agendamento realizado com sucesso!", "ok")
        else:
            msg("\n⚠ Opção inválida", "warn")

    elif escolha == "4":
        return
    else:
        msg("\n⚠ Opção inválida", "warn")

def limpar_dados(usuario_logado: str):
    if usuario_logado != "admin":
        msg("\n⚠ Apenas o ADMIN pode acessar esta opção.", "warn")
        return
    print("\n=== MENU DE LIMPEZA DE DADOS ===")
    print("1 - Limpar relatórios")
    print("2 - Limpar agendamentos")
    print("3 - Limpar alunos")
    print("4 - Limpar tudo")
    print("5 - Voltar")
    escolha = input("Escolha uma opção: ").strip()
    if escolha == "1":
        msg("\n🧹 Limpando relatórios...", "info")
        time.sleep(0.6)
        for arq in (ARQ_REL, ARQ_REL_XLSX):
            try:
                if arq.exists(): arq.unlink()
            except Exception:
                pass
        msg("\n✅ Relatórios apagados com sucesso!", "ok")
    elif escolha == "2":
        msg("\n🧹 Limpando agendamentos...", "info")
        time.sleep(0.6)
        try:
            if ARQ_AG.exists(): ARQ_AG.unlink()
            xlsx = Path("agendamentos.xlsx")
            if xlsx.exists(): xlsx.unlink()
        except Exception:
            pass
        msg("\n✅ Agendamentos apagados com sucesso!", "ok")
    elif escolha == "3":
        msg("\n🧹 Limpando alunos...", "info")
        time.sleep(0.6)
        for arq in (ARQ_ALUNOS, ARQ_ALUNOS_XLSX):
            try:
                if arq.exists(): arq.unlink()
            except Exception:
                pass
        msg("\n✅ Alunos apagados com sucesso!", "ok")
    elif escolha == "4":
        msg("\n🧹 Limpando todos os dados do sistema...", "info")
        time.sleep(0.6)
        arquivos = [ARQ_REL, ARQ_REL_XLSX, ARQ_AG, ARQ_ALUNOS, ARQ_ALUNOS_XLSX, Path("agendamentos.xlsx")]
        for arq in arquivos:
            try:
                if arq.exists(): arq.unlink()
            except Exception:
                pass
        msg("\n✅ Todos os dados foram apagados com sucesso!", "ok")
    elif escolha == "5":
        return
    else:
        msg("\n⚠ Opção inválida.", "warn")

def login():
    acesso_liberado = False
    usuario_logado = None
    while not acesso_liberado:
        print("\n=== LOGIN ===")
        usuario = input("Digite seu login: ").lower().strip()
        senha = pwinput.pwinput(prompt="Digite a senha: ", mask="*").strip()
        limpar_tela()
        if usuario in USERS and USERS[usuario]["senha"] == senha:
            msg("\n✅ Login efetuado com sucesso", "ok")
            usuario_logado = usuario
            acesso_liberado = True
            time.sleep(0.4)
        else:
            msg("\n❌ Login inválido. Tente novamente", "err")
            time.sleep(0.6)
            limpar_tela()
    return usuario_logado

def main():
    usuario = login()
    while True:
        print("\n=== MENU PRINCIPAL ===")
        print("1 - Computadores")
        print("2 - Agendamento")
        print("3 - Relatório")
        print("4 - Sair")
        if usuario == "admin":
            print("5 - Limpar dados")
        opcao = input("Escolha o que deseja fazer: ").strip()
        if opcao == "1":
            menu_computadores(usuario)
        elif opcao == "2":
            menu_agendamento(usuario)
        elif opcao == "3":
            gerar_relatorio(usuario)
        elif opcao == "4":
            msg("\n👋 Saindo do sistema...", "info")
            time.sleep(0.4)
            break
        elif opcao == "5" and usuario == "admin":
            limpar_dados(usuario)
        else:
            msg("\n⚠ Opção inválida.", "warn")

if __name__ == "__main__":
    main()
