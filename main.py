import flet as ft
import openpyxl
import os

# Banco de dados simulado na memória (Em produção, pode ser salvo em arquivo)
USUARIOS = {"admin": "1234"}
PRODUTOS_CADASTRADOS = {}  # {codigo: {"nome": nome, "lastro_cx": cx, "pallet_cx": cx}}
INVENTARIO_ATUAL = {}  # {codigo: {"pallets": 0, "lastros": 0, "caixas": 0, "unidades": 0}}


def main(page: ft.Page):
    page.title = "DISF_DPO"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.scroll = ft.ScrollMode.AUTO
    page.window_width = 450
    page.window_height = 750

    usuario_logado = ""

    # --- FUNÇÕES DE NAVEGAÇÃO E LÓGICA ---
    def ir_para_login(e=None):
        page.clean()
        page.add(tela_login())

    def ir_para_menu(e=None):
        page.clean()
        page.add(tela_menu_principal())

    def ir_para_cadastro(e=None):
        page.clean()
        page.add(tela_cadastro_qr())

    def ir_para_inventario(e=None):
        page.clean()
        page.add(tela_iniciar_inventario())

    # --- COMPONENTES VISUAIS REUTILIZÁVEIS ---
    def obter_rodape():
        return ft.Container(
            content=ft.Text(
                "Criado por: Elton Temoteo (88) 99798-2788",
                size=12,
                color=ft.Colors.GREY_600,
                weight=ft.FontWeight.BOLD,
                text_align=ft.TextAlign.CENTER
            ),
            alignment=ft.alignment.center,
            margin=ft.margin.only(top=20)
        )

    def obter_logo():
        return ft.Container(
            content=ft.Icon(ft.Icons.LOCAL_SHIPPING, size=60, color=ft.Colors.BLUE_800),
            alignment=ft.alignment.center,
            margin=ft.margin.only(bottom=10)
        )

    # --- 1. TELA DE LOGIN E CRIAÇÃO DE USUÁRIOS ---
    def tela_login():
        txt_user = ft.TextField(label="Usuário", width=300)
        txt_pass = ft.TextField(label="Senha", password=True, can_reveal_password=True, width=300)
        txt_erro = ft.Text("", color=ft.Colors.RED)

        def realizar_login(e):
            nonlocal usuario_logado
            u, p = txt_user.value.strip(), txt_pass.value.strip()
            if u in USUARIOS and USUARIOS[u] == p:
                usuario_logado = u
                ir_para_menu()
            else:
                txt_erro.value = "Usuário ou senha incorretos."
                page.update()

        def criar_usuario(e):
            u, p = txt_user.value.strip(), txt_pass.value.strip()
            if usuario_logado != "admin":
                txt_erro.value = "Apenas o admin pode criar novos usuários!"
            elif u and p:
                USUARIOS[u] = p
                txt_erro.value = f"Usuário '{u}' criado com sucesso!"
                txt_erro.color = ft.Colors.GREEN_700
            else:
                txt_erro.value = "Preencha usuário e senha para criar."
            page.update()

        return ft.Column(
            controls=[
                obter_logo(),
                ft.Text("DISF_DPO - Login", size=22, weight=ft.FontWeight.BOLD),
                txt_user, txt_pass, txt_erro,
                ft.ElevatedButton("Entrar", on_click=realizer_login, width=300, bgcolor=ft.Colors.BLUE_800,
                                  color=ft.Colors.WHITE),
                ft.TextButton("Criar Novo Usuário (Admin)", on_click=criar_usuario),
                obter_rodape()
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

    # --- 2. TELA MENU PRINCIPAL ---
    def tela_menu_principal():
        def zerar_relatorio(e):
            INVENTARIO_ATUAL.clear()
            dlg = ft.AlertDialog(title=ft.Text("Relatório zerado com sucesso!"))
            page.open(dlg)

        def gerar_excel(e):
            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Inventário"
                ws.append(["Código/QR", "Produto", "Pallets", "Lastros", "Caixas", "Unidades"])
                for cod, qtds in INVENTARIO_ATUAL.items():
                    nome = PRODUTOS_CADASTRADOS.get(cod, {}).get("nome", "Desconhecido")
                    ws.append([cod, nome, qtds["pallets"], qtds["lastros"], qtds["caixas"], qtds["unidades"]])
                wb.save("relatorio_inventario.xlsx")
                page.open(ft.AlertDialog(title=ft.Text("Excel Gerado com Sucesso!")))
            except Exception as err:
                page.open(ft.AlertDialog(title=ft.Text(f"Erro: {str(err)}")))

        return ft.Column(
            controls=[
                obter_logo(),
                ft.Text("MENU PRINCIPAL", size=20, weight=ft.FontWeight.BOLD),
                ft.ElevatedButton("🔄 ZERAR RELATÓRIO", on_click=zerar_relatorio, bgcolor=ft.Colors.RED_700,
                                  color=ft.Colors.WHITE, width=300),
                ft.Divider(height=20),
                ft.ElevatedButton("01. CADASTRAR QR OU CÓDIGO", on_click=ir_para_cadastro, width=300),
                ft.ElevatedButton("02. INICIAR INVENTÁRIO", on_click=ir_para_inventario, width=300),
                ft.ElevatedButton("03. GERAR EXCEL", on_click=gerar_excel, width=300),
                ft.ElevatedButton("04. SAIR", on_click=lambda _: page.window_destroy(), width=300,
                                  bgcolor=ft.Colors.GREY_400),
                obter_rodape()
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

    # --- 3. TELA 01: CADASTRAR/MODIFICAR QR OU CÓDIGO ---
    def tela_cadastro_qr():
        txt_cod = ft.TextField(label="Código de Barras ou QR Code", width=300)
        txt_nome = ft.TextField(label="Nome do Produto (Ex: Skol 600ml)", width=300)
        txt_status = ft.Text("")

        def salvar_produto(e):
            c, n = txt_cod.value.strip(), txt_nome.value.strip()
            if c and n:
                PRODUTOS_CADASTRADOS[c] = {"nome": n}
                txt_status.value = f"✅ Produto salvo/modificado!"
                txt_status.color = ft.Colors.GREEN_700
                txt_cod.value = ""
                txt_nome.value = ""
            else:
                txt_status.value = "⚠️ Preencha todos os campos."
                txt_status.color = ft.Colors.ORANGE_700
            page.update()

        return ft.Column(
            controls=[
                ft.Text("Cadastrar/Trocar QR ou Código", size=18, weight=ft.FontWeight.BOLD),
                txt_cod, txt_nome, txt_status,
                ft.ElevatedButton("Salvar Produto", on_click=salvar_produto, width=300),
                ft.TextButton("⬅️ Voltar para Tela Inicial", on_click=ir_para_menu),
                obter_rodape()
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

    # --- 4. TELA 02: INICIAR INVENTÁRIO (SOMA ACUMULADA) ---
    def tela_iniciar_inventario():
        txt_busca = ft.TextField(label="Ler QR ou Inserir Código", width=300)
        txt_prod_nome = ft.Text("Produto: Selecione um código", size=16, weight=ft.FontWeight.BOLD,
                                color=ft.Colors.BLUE_800)

        # Campos de Logística
        input_pallet = ft.TextField(label="Pallet", value="0", width=65, text_align=ft.TextAlign.CENTER)
        input_lastro = ft.TextField(label="Lastro", value="0", width=65, text_align=ft.TextAlign.CENTER)
        input_caixa = ft.TextField(label="Caixa", value="0", width=65, text_align=ft.TextAlign.CENTER)
        input_unidade = ft.TextField(label="Unidade", value="0", width=65, text_align=ft.TextAlign.CENTER)

        txt_status = ft.Text("")

        def buscar_produto(e):
            cod = txt_busca.value.strip()
            if cod in PRODUTOS_CADASTRADOS:
                txt_prod_nome.value = f"Produto: {PRODUTOS_CADASTRADOS[cod]['nome']}"
            else:
                txt_prod_nome.value = "❌ Produto não cadastrado!"
            page.update()

        txt_busca.on_change = buscar_produto

        def salvar_contagem(e):
            cod = txt_busca.value.strip()
            if cod not in PRODUTOS_CADASTRADOS:
                txt_status.value = "❌ Não é possível salvar um produto não cadastrado."
                txt_status.color = ft.Colors.RED_700
                page.update()
                return

            # Captura os valores digitados (converte para 0 se estiver vazio)
            p = int(input_pallet.value) if input_pallet.value.isdigit() else 0
            l = int(input_lastro.value) if input_lastro.value.isdigit() else 0
            c = int(input_caixa.value) if input_caixa.value.isdigit() else 0
            u = int(input_unidade.value) if input_unidade.value.isdigit() else 0

            # Lógica inteligente: Se já existia contagem desse produto em outro lugar, ele SOMA
            if cod in INVENTARIO_ATUAL:
                INVENTARIO_ATUAL[cod]["pallets"] += p
                INVENTARIO_ATUAL[cod]["lastros"] += l
                INVENTARIO_ATUAL[cod]["caixas"] += c
                INVENTARIO_ATUAL[cod]["unidades"] += u
            else:
                INVENTARIO_ATUAL[cod] = {"pallets": p, "lastros": l, "caixas": c, "unidades": u}

            txt_status.value = "💾 Contagem acumulada com sucesso!"
            txt_status.color = ft.Colors.GREEN_700

            # Limpa os campos para o próximo produto
            txt_busca.value = ""
            txt_prod_nome.value = "Produto: Selecione um código"
            input_pallet.value = "0"
            input_lastro.value = "0"
            input_caixa.value = "0"
            input_unidade.value = "0"
            page.update()

        return ft.Column(
            controls=[
                ft.Text("Controle de Inventário", size=18, weight=ft.FontWeight.BOLD),
                txt_busca,
                ft.Container(content=txt_prod_nome, padding=10, bgcolor=ft.Colors.GREY_200, borderRadius=5, width=300),
                ft.Text("Quantidades encontradas:", size=14),
                ft.Row([input_pallet, input_lastro, input_caixa, input_unidade], alignment=ft.MainAxisAlignment.CENTER),
                txt_status,
                ft.ElevatedButton("Salvar e Próximo", on_click=salvar_contagem, bgcolor=ft.Colors.GREEN_700,
                                  color=ft.Colors.WHITE, width=300),
                ft.TextButton("⬅️ Voltar para Tela Inicial", on_click=ir_para_menu),
                obter_rodape()
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

    # Inicia o app direto na tela de login
    ir_para_login()


if __name__ == "__main__":
    ft.app(target=main)