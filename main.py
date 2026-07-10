import flet as ft
import openpyxl
import os
import csv

# --- CONFIGURAÇÕES DO CSV (AJUSTE AQUI SE MUDAR A PLANILHA) ---
NOME_ARQUIVO_PRODUTOS = "PRODUTOS.csv"
SEPARADOR_CSV = ';'  # Se usar vírgula, troque para ','
PULAR_CABECALHO = True  # Pular a primeira linha (CODIGO;PRODUTO...)
COLUNA_NOME_PRODUTO = 1  # 0 é Código, 1 é Produto

# --- BANCO DE DADOS DA DISTRIBUIDORA (Carregado do CSV) ---
LISTA_PRODUTOS_OFICIAL = []
PRODUTOS_CSV = {}  # {codigo: nome}


def carregar_produtos_csv():
    global LISTA_PRODUTOS_OFICIAL, PRODUTOS_CSV
    LISTA_PRODUTOS_OFICIAL.clear()
    PRODUTOS_CSV.clear()

    if os.path.exists(NOME_ARQUIVO_PRODUTOS):
        # focado no padrão cp1252 que é o formato nativo do Excel no Windows brasileiro
        encodings_para_tentar = ['cp1252', 'utf-8-sig', 'iso-8859-1', 'utf-8']

        for enc in encodings_para_tentar:
            try:
                with open(NOME_ARQUIVO_PRODUTOS, mode='r', encoding=enc) as f:
                    conteudo = f.read()
                    # Detecta o separador real
                    separador_real = ';' if ';' in conteudo else ','
                    f.seek(0)

                    leitor = csv.reader(f, delimiter=separador_real)

                    if PULAR_CABECALHO:
                        next(leitor, None)  # Pula o cabeçalho

                    for linha in leitor:
                        if linha and len(linha) > COLUNA_NOME_PRODUTO:
                            # Pega a coluna correta (PRODUTO)
                            produto_codigo = linha[0].strip()
                            produto_nome = linha[COLUNA_NOME_PRODUTO].strip()
                            if produto_nome:
                                LISTA_PRODUTOS_OFICIAL.append(produto_nome)
                                PRODUTOS_CSV[produto_codigo] = produto_nome

                if LISTA_PRODUTOS_OFICIAL:
                    # Ordena a lista para facilitar a busca
                    LISTA_PRODUTOS_OFICIAL.sort()
                    break
            except Exception:
                continue

        if not LISTA_PRODUTOS_OFICIAL:
            LISTA_PRODUTOS_OFICIAL = ["Erro: Não foi possível processar as linhas do CSV."]
    else:
        LISTA_PRODUTOS_OFICIAL = ["Arquivo PRODUTOS.csv não encontrado na pasta."]


# Inicializa a lista de produtos
carregar_produtos_csv()

# Banco de dados simulado na memória para o inventário
USUARIOS = {"disfrota": "1234"}
PRODUTOS_CADASTRADOS = {}  # {codigo: {"nome": nome}}
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
                color=ft.colors.GREY_600,
                weight="bold",
                text_align=ft.TextAlign.CENTER
            ),
            alignment=ft.alignment.center,
            margin=ft.margin.only(top=20)
        )

    def obter_logo():
        return ft.Container(
            content=ft.Icon(ft.icons.LOCAL_SHIPPING, size=60, color=ft.colors.BLUE_800),
            alignment=ft.alignment.center,
            margin=ft.margin.only(bottom=10)
        )

    # --- 1. TELA DE LOGIN ---
    def tela_login():
        txt_user = ft.TextField(label="Usuário", width=300)
        txt_pass = ft.TextField(label="Senha", password=True, can_reveal_password=True, width=300)
        txt_erro = ft.Text("", color=ft.colors.RED)

        def realizar_login(e):
            nonlocal usuario_logado
            u, p = txt_user.value.strip(), txt_pass.value.strip()
            if u in USUARIOS and USUARIOS[u] == p:
                usuario_logado = u
                ir_para_menu()
            else:
                txt_erro.value = "Usuário ou senha incorretos."
                page.update()

        return ft.Column(
            controls=[
                obter_logo(),
                ft.Text("DISF_DPO - Login", size=22, weight="bold"),
                txt_user, txt_pass, txt_erro,
                ft.ElevatedButton("Entrar", on_click=realizar_login, width=300, bgcolor=ft.colors.BLUE_800,
                                  color=ft.colors.WHITE),
                obter_rodape()
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

    # --- 2. TELA MENU PRINCIPAL ---
    def tela_menu_principal():
        def zerar_relatorio(e):
            INVENTARIO_ATUAL.clear()
            page.open(ft.AlertDialog(title=ft.Text("Relatório zerado com sucesso!")))

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
                ft.Text("MENU PRINCIPAL", size=20, weight="bold"),
                ft.ElevatedButton("🔄 ZERAR RELATÓRIO", on_click=zerar_relatorio, bgcolor=ft.colors.RED_700,
                                  color=ft.colors.WHITE, width=300),
                ft.Divider(height=20),
                ft.ElevatedButton("01. CADASTRAR QR OU CÓDIGO", on_click=ir_para_cadastro, width=300),
                ft.ElevatedButton("02. INICIAR INVENTÁRIO", on_click=ir_para_inventario, width=300),
                ft.ElevatedButton("03. GERAR EXCEL", on_click=gerar_excel, width=300),
                ft.ElevatedButton("04. SAIR", on_click=lambda _: page.window_destroy(), width=300,
                                  bgcolor=ft.colors.GREY_400),
                obter_rodape()
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

    # --- 3. TELA 01: CADASTRAR QR COM BUSCA DA PLANILHA E CÂMERA ---
    def tela_cadastro_qr():
        txt_cod = ft.TextField(label="Código de Barras ou QR Code", width=220)

        def abrir_camera_cadastro(e):
            # No PC simula, no APK final usaremos o leitor de câmera nativo
            txt_cod.value = "7891991000853"  # Exemplo: Skol Litrão
            txt_status.value = "📸 Código capturado pela câmera!"
            txt_status.color = ft.colors.BLUE_700
            page.update()

        btn_camera = ft.IconButton(icon=ft.icons.CAMERA_ALT, on_click=abrir_camera_cadastro,
                                   icon_color=ft.colors.BLUE_800, tooltip="Abrir Câmera")

        txt_busca_prod = ft.TextField(label="Digite o nome do produto...", width=300)
        txt_selecionado = ft.Text("Nenhum produto selecionado", color=ft.colors.BLUE_800, weight="bold")
        lista_opcoes = ft.ListView(expand=1, spacing=5, height=120, width=300)

        txt_status = ft.Text("")
        produto_escolhido = ""

        def filtrar_produtos(e):
            nonlocal produto_escolhido
            termo = txt_busca_prod.value.lower().strip()
            novos_controles = []

            if termo:
                # Varre os produtos reais vindos do seu CSV
                for prod in LISTA_PRODUTOS_OFICIAL:
                    if termo in prod.lower():
                        # Cria uma função separada para capturar o nome correto do loop
                        def criar_evento_clique(nome_item):
                            return lambda _: selecionar_opcao(nome_item)

                        def selecionar_opcao(nome_prod):
                            nonlocal produto_escolhido
                            produto_escolhido = nome_prod
                            txt_selecionado.value = f"✓ Selecionado: {nome_prod}"
                            txt_busca_prod.value = nome_prod
                            lista_opcoes.controls = []  # Limpa as sugestões após clicar
                            page.update()

                        novos_controles.append(
                            ft.ListTile(
                                title=ft.Text(prod, size=14),
                                leading=ft.Icon(ft.icons.LOCAL_DRINK, size=20),
                                on_click=criar_evento_clique(prod)
                            )
                        )
            # Aplica os novos itens de forma nativa do Flet
            lista_opcoes.controls = novos_controles
            page.update()

        txt_busca_prod.on_change = filtrar_produtos

        def salvar_produto(e):
            nonlocal produto_escolhido
            c = txt_cod.value.strip()
            if c and produto_escolhido:
                PRODUTOS_CADASTRADOS[c] = {"nome": produto_escolhido}
                txt_status.value = f"✅ Vinculado com sucesso!"
                txt_status.color = ft.colors.GREEN_700

                # Reseta tudo para o próximo cadastro
                txt_cod.value = ""
                txt_busca_prod.value = ""
                txt_selecionado.value = "Nenhum produto selecionado"
                produto_escolhido = ""
                lista_opcoes.controls = []
            else:
                txt_status.value = "⚠️ Leia o QR/Código e selecione um produto da lista."
                txt_status.color = ft.colors.ORANGE_700
            page.update()

        return ft.Column(
            controls=[
                ft.Text("Cadastrar QR ou Código", size=18, weight="bold"),
                ft.ElevatedButton("Salvar Vínculo e Próximo", on_click=salvar_produto, width=300,
                                  bgcolor=ft.colors.GREEN_700, color=ft.colors.WHITE),
                ft.ElevatedButton("Voltar ao Menu Inicial", on_click=ir_para_menu, width=300,
                                  bgcolor=ft.colors.BLUE_800, color=ft.colors.WHITE),
                ft.Row([txt_cod, btn_camera], alignment=ft.MainAxisAlignment.CENTER, spacing=10),
                txt_busca_prod,
                lista_opcoes,  # Janelinha interna que mostra as opções dinâmicas
                ft.Container(content=txt_selecionado, padding=8, bgcolor=ft.colors.BLUE_50, border_radius=5, width=300),
                txt_status,

                obter_rodape()
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

    # --- 4. TELA 02: INICIAR INVENTÁRIO COM BOTÃO DE CÂMERA ---
    def tela_iniciar_inventario():
        txt_busca = ft.TextField(label="Ler QR ou Inserir Código", width=220)

        def abrir_camera_inventario(e):
            txt_busca.value = "7891991000853"  # Simula Skol Litrão vinculada
            buscar_produto(None)
            page.update()

        btn_camera_inv = ft.IconButton(icon=ft.icons.CAMERA_ALT, on_click=abrir_camera_inventario,
                                       icon_color=ft.colors.BLUE_800)

        txt_prod_nome = ft.Text("Produto: Selecione um código", size=16, weight="bold", color=ft.colors.BLUE_800)

        # Campos de Logística na ordem solicitada
        input_pallet = ft.TextField(label="Pallet", value="0", width=65, text_align=ft.TextAlign.CENTER)
        input_lastro = ft.TextField(label="Lastro", value="0", width=65, text_align=ft.TextAlign.CENTER)
        input_caixa = ft.TextField(label="Caixa", value="0", width=65, text_align=ft.TextAlign.CENTER)
        input_unidade = ft.TextField(label="Unidade", value="0", width=65, text_align=ft.TextAlign.CENTER)

        txt_status = ft.Text("")

        def buscar_produto(e):
            cod = txt_busca.value.strip()
            if cod in PRODUTOS_CADASTRADOS:
                txt_prod_nome.value = f"Produto: {PRODUTOS_CADASTRADOS[cod]['nome']}"
            elif cod in PRODUTOS_CSV:
                txt_prod_nome.value = f"Código: {cod} | Produto: {PRODUTOS_CSV[cod]}"
            else:
                txt_prod_nome.value = "❌ Produto não encontrado!"
            page.update()

        txt_busca.on_change = buscar_produto

        def salvar_contagem(e):
            cod = txt_busca.value.strip()
            if cod not in PRODUTOS_CADASTRADOS and cod not in PRODUTOS_CSV:
                txt_status.value = "❌ Produto não encontrado."
                txt_status.color = ft.colors.RED_700
                page.update()
                return

            # Validação para garantir que são números
            p = int(input_pallet.value) if input_pallet.value.isdigit() else 0
            l = int(input_lastro.value) if input_lastro.value.isdigit() else 0
            c = int(input_caixa.value) if input_caixa.value.isdigit() else 0
            u = int(input_unidade.value) if input_unidade.value.isdigit() else 0

            # Lógica de acúmulo inteligente
            if cod in INVENTARIO_ATUAL:
                INVENTARIO_ATUAL[cod]["pallets"] += p
                INVENTARIO_ATUAL[cod]["lastros"] += l
                INVENTARIO_ATUAL[cod]["caixas"] += c
                INVENTARIO_ATUAL[cod]["unidades"] += u
            else:
                INVENTARIO_ATUAL[cod] = {"pallets": p, "lastros": l, "caixas": c, "unidades": u}

            if cod in PRODUTOS_CSV and cod not in PRODUTOS_CADASTRADOS:
                PRODUTOS_CADASTRADOS[cod] = {"nome": PRODUTOS_CSV[cod]}

            txt_status.value = "💾 Contagem acumulada com sucesso!"
            txt_status.color = ft.colors.GREEN_700

            # Limpa campos para o próximo item
            txt_busca.value = ""
            txt_prod_nome.value = "Produto: Selecione um código"
            input_pallet.value = "0"
            input_lastro.value = "0"
            input_caixa.value = "0"
            input_unidade.value = "0"
            page.update()

        return ft.Column(
            controls=[
                ft.Text("Controle de Inventário", size=18, weight="bold"),
                ft.Row([txt_busca, btn_camera_inv], alignment=ft.MainAxisAlignment.CENTER, spacing=10),
                ft.Container(content=txt_prod_nome, padding=10, bgcolor=ft.colors.GREY_200, border_radius=5, width=300),
                ft.Text("Quantidades encontradas:", size=14),
                ft.Row([input_pallet, input_lastro, input_caixa, input_unidade], alignment=ft.MainAxisAlignment.CENTER),
                txt_status,
                ft.ElevatedButton("Salvar Contagem e Próximo", on_click=salvar_contagem, bgcolor=ft.colors.GREEN_700,
                                  color=ft.colors.WHITE, width=300),
                ft.ElevatedButton("Voltar ao Menu Inicial", on_click=ir_para_menu, width=300,
                                  bgcolor=ft.colors.BLUE_800, color=ft.colors.WHITE),
                obter_rodape()
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

    # Inicia o app na tela de login
    ir_para_login()


if __name__ == "__main__":
    ft.app(target=main)
