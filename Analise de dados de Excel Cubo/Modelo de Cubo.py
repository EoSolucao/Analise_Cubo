import flet as ft
import pandas as pd
import os
from typing import Dict, List, Any


def main(page: ft.Page):
    page.title = "Análise de Dados Excel Múltiplos"
    page.theme_mode = "light"
    page.window_width = 1600
    page.window_height = 900

    dfs: Dict[str, pd.DataFrame] = {}
    campos: Dict[str, List[str]] = {"linha": [], "valor": [], "filtro": []}
    join_keys: Dict[str, Any] = {"left": None, "right": None}
    filtros: Dict[str, ft.Dropdown] = {}

    def show_snackbar(message: str, color: str = "green"):
        page.snack_bar = ft.SnackBar(content=ft.Text(message), bgcolor=color)
        page.snack_bar.open = True
        page.update()

    def criar_drop_target(texto: str, cor: str, key: str) -> ft.DragTarget:
        list_view = ft.ListView(expand=1, spacing=5, padding=10, auto_scroll=True)
        return ft.DragTarget(
            group="campos",
            content=ft.Container(
                width=300,
                height=100,
                bgcolor=cor,
                border_radius=10,
                content=ft.Column([ft.Text(f"Campo {texto.capitalize()}"), list_view]),
                alignment=ft.alignment.center,
            ),
            on_accept=lambda e: adicionar_campo(e, key, list_view),
        )

    def carregar_excel(e: ft.FilePickerResultEvent):
        if e.files:
            for file in e.files:
                df = pd.read_excel(file.path)
                dfs[file.name] = df
            atualizar_lista_arquivos()
            atualizar_campos_disponiveis()
            atualizar_cubo()
            show_snackbar("Arquivos carregados com sucesso!")

    def atualizar_lista_arquivos():
        lista_arquivos.controls = [
            ft.Row(
                [
                    ft.Text(f"{nome} ({df.shape[0]} linhas, {df.shape[1]} colunas)"),
                    ft.IconButton(
                        ft.icons.DELETE, on_click=lambda _, n=nome: remover_arquivo(n)
                    ),
                ]
            )
            for nome, df in dfs.items()
        ]
        page.update()

    def remover_arquivo(nome: str):
        del dfs[nome]
        atualizar_lista_arquivos()
        atualizar_campos_disponiveis()
        atualizar_cubo()

    def atualizar_campos_disponiveis():
        colunas = set()
        for df in dfs.values():
            colunas.update(df.columns)
        drags.controls = [criar_drag(coluna) for coluna in sorted(colunas)]
        page.update()

    def criar_drag(texto: str) -> ft.Draggable:
        return ft.Draggable(
            group="campos",
            content=ft.Container(
                content=ft.Text(texto),
                padding=10,
                bgcolor=ft.colors.BLUE_50,
                border_radius=5,
            ),
            data=texto,
        )

    def adicionar_campo(e, key: str, list_view: ft.ListView):
        src = page.get_control(e.src_id)
        campo = src.data
        if campo not in campos[key]:
            campos[key].append(campo)
            list_view.controls.append(
                ft.Row(
                    [
                        ft.Text(campo),
                        ft.IconButton(
                            ft.icons.CLOSE,
                            on_click=lambda _, c=campo, k=key, lv=list_view: remover_campo(
                                c, k, lv
                            ),
                        ),
                    ]
                )
            )
            if key == "filtro":
                criar_filtro(campo)
            page.update()
            atualizar_cubo()

    def remover_campo(campo: str, key: str, list_view: ft.ListView):
        campos[key].remove(campo)
        list_view.controls = [
            c for c in list_view.controls if c.controls[0].value != campo
        ]
        if key == "filtro":
            remover_filtro(campo)
        page.update()
        atualizar_cubo()

    def criar_filtro(campo: str):
        valores_unicos = set()
        for df in dfs.values():
            if campo in df.columns:
                valores_unicos.update(df[campo].astype(str).unique())
        dropdown = ft.Dropdown(
            label=f"Filtrar {campo}",
            options=[ft.dropdown.Option(valor) for valor in sorted(valores_unicos)],
            on_change=lambda _: aplicar_filtros(),
        )
        filtros[campo] = dropdown
        area_filtros.controls.append(ft.Row([ft.Text(campo), dropdown]))
        page.update()

    def remover_filtro(campo: str):
        if campo in filtros:
            del filtros[campo]
            area_filtros.controls = [
                c for c in area_filtros.controls if c.controls[0].value != campo
            ]
            page.update()

    def aplicar_filtros():
        atualizar_cubo()

    def atualizar_cubo():
        if not dfs or not any(campos.values()):
            resultado.columns = [ft.DataColumn(ft.Text("Sem dados"))]
            resultado.rows = []
            page.update()
            return

        df_final = None
        for i, (df_name, df) in enumerate(dfs.items()):
            if i == 0:
                df_final = df
            else:
                if join_keys["left"] and join_keys["right"]:
                    try:
                        df_final = pd.merge(
                            df_final,
                            df,
                            how=tipo_join.value,
                            left_on=join_keys["left"],
                            right_on=join_keys["right"],
                        )
                    except KeyError as e:
                        show_snackbar(f"Erro ao juntar tabelas: {e}", "red")
                        return
                else:
                    show_snackbar(
                        f"Aviso: Não foi possível juntar a tabela {df_name}", "yellow"
                    )

        for campo, dropdown in filtros.items():
            if dropdown.value:
                df_final = df_final[df_final[campo].astype(str) == dropdown.value]

        colunas_selecionadas = campos["linha"] + campos["valor"] + campos["filtro"]
        colunas_selecionadas = [
            col for col in colunas_selecionadas if col in df_final.columns
        ]

        if not colunas_selecionadas:
            colunas_selecionadas = df_final.columns.tolist()[:1]

        df_final = df_final[colunas_selecionadas]

        if campos["valor"] and campos["linha"]:
            try:
                df_final = df_final.groupby(campos["linha"], as_index=False).agg(
                    {col: calcular_valor for col in campos["valor"]}
                )
            except ValueError as e:
                show_snackbar(f"Erro ao agrupar dados: {e}", "red")
                return

        resultado.columns = [
            ft.DataColumn(ft.Text(col, weight="bold")) for col in df_final.columns
        ]
        resultado.rows = [
            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(formatar_valor(value, col)))
                    for col, value in row.items()
                ]
            )
            for _, row in df_final.iterrows()
        ]

        page.update()

    def calcular_valor(series):
        if operacao_valor.value == "Soma":
            return series.sum()
        elif operacao_valor.value == "Contagem":
            return series.count()
        elif operacao_valor.value == "Máximo":
            return series.max()
        elif operacao_valor.value == "Mínimo":
            return series.min()
        elif operacao_valor.value == "Média":
            return series.mean()
        return series

    def formatar_valor(value, column):
        if pd.isna(value):
            return ""

        if isinstance(value, (int, float)):
            if formatacao_valor.value == "Número":
                return f"{value:,.2f}"
            elif formatacao_valor.value == "Inteiro":
                return f"{int(value):,}"
        elif isinstance(value, str):
            if formatacao_valor.value == "Data":
                try:
                    return pd.to_datetime(value).strftime("%d/%m/%Y")
                except:
                    return value
            elif formatacao_valor.value == "Hora":
                try:
                    return pd.to_datetime(value).strftime("%H:%M")
                except:
                    return value

        return str(value)

    def adicionar_campo_juncao(e: ft.DragTargetAcceptEvent, side: str):
        src = page.get_control(e.src_id)
        campo = src.data
        join_keys[side] = campo
        join_areas[side].content.content.controls[1].controls = [
            ft.Row(
                [
                    ft.Text(campo),
                    ft.IconButton(
                        ft.icons.CLOSE,
                        on_click=lambda _, s=side: remover_campo_juncao(s),
                    ),
                ]
            )
        ]
        page.update()
        atualizar_cubo()

    def remover_campo_juncao(side: str):
        join_keys[side] = None
        join_areas[side].content.content.controls[1].controls = []
        page.update()
        atualizar_cubo()

    def salvar_excel(e):
        if not resultado.rows:
            show_snackbar("Não há dados para salvar.", "red")
            return

        base_path = os.path.dirname(os.path.abspath(__file__))
        i = 1
        while True:
            save_path = os.path.join(base_path, f"resultado_{i}.xlsx")
            if not os.path.exists(save_path):
                break
            i += 1

        df_final = pd.DataFrame(
            [[cell.content.value for cell in row.cells] for row in resultado.rows],
            columns=[col.label for col in resultado.columns],
        )

        df_final.to_excel(save_path, index=False, sheet_name="Resultado")
        show_snackbar(f"Arquivo salvo com sucesso: {save_path}")

    file_picker = ft.FilePicker(on_result=carregar_excel)
    page.overlay.append(file_picker)

    btn_selecionar_arquivos = ft.ElevatedButton(
        "Selecionar arquivos Excel",
        icon=ft.icons.UPLOAD_FILE,
        on_click=lambda _: file_picker.pick_files(
            allow_multiple=True, allowed_extensions=["xlsx"]
        ),
    )

    lista_arquivos = ft.Column()

    tipo_join = ft.Dropdown(
        label="Tipo de Join",
        options=[
            ft.dropdown.Option("inner", "Inner Join"),
            ft.dropdown.Option("left", "Left Join"),
            ft.dropdown.Option("right", "Right Join"),
            ft.dropdown.Option("outer", "Outer Join"),
        ],
        value="inner",
        on_change=lambda _: atualizar_cubo(),
    )

    join_info_icon = ft.IconButton(
        icon=ft.icons.INFO,
        tooltip="Informações sobre tipos de Join",
        on_click=lambda _: show_join_info(),
    )

    def show_join_info():
        page.dialog = ft.AlertDialog(
            title=ft.Text("Tipos de Join"),
            content=ft.Text(
                "- Inner Join: Retorna apenas as linhas que têm correspondência em ambas as tabelas.\n"
                "- Left Join: Retorna todas as linhas da tabela à esquerda e as linhas correspondentes da tabela à direita.\n"
                "- Right Join: Retorna todas as linhas da tabela à direita e as linhas correspondentes da tabela à esquerda.\n"
                "- Outer Join: Retorna todas as linhas quando há uma correspondência em uma das tabelas."
            ),
            actions=[ft.TextButton("Fechar", on_click=lambda _: close_dialog())],
        )
        page.dialog.open = True
        page.update()

    def close_dialog():
        page.dialog.open = False
        page.update()

    drags = ft.Row([], wrap=True, spacing=10)
    campos_text = {key: ft.Column() for key in campos}
    drop_linha = criar_drop_target("Linha", ft.colors.BLUE_100, "linha")
    drop_valor = criar_drop_target("Valor", ft.colors.GREEN_100, "valor")
    drop_filtro = criar_drop_target("Filtro", ft.colors.PURPLE_100, "filtro")

    def criar_join_area(texto: str, side: str) -> ft.DragTarget:
        return ft.DragTarget(
            group="campos",
            content=ft.Container(
                width=300,
                height=100,
                bgcolor=ft.colors.TEAL_100,
                border_radius=10,
                content=ft.Column(
                    [
                        ft.Text(f"Junção {texto}"),
                        ft.ListView(expand=1, spacing=5, padding=10, auto_scroll=True),
                    ]
                ),
                alignment=ft.alignment.center,
            ),
            on_accept=lambda e: adicionar_campo_juncao(e, side),
        )

    join_areas = {
        "left": criar_join_area("1 (Esquerda)", "left"),
        "right": criar_join_area("2 (Direita)", "right"),
    }

    area_filtros = ft.Column()

    resultado = ft.DataTable(
        columns=[ft.DataColumn(ft.Text("Sem dados"))],
        rows=[],
        border=ft.border.all(2, ft.colors.GREY_400),
        border_radius=10,
        vertical_lines=ft.border.BorderSide(1, ft.colors.GREY_400),
        horizontal_lines=ft.border.BorderSide(1, ft.colors.GREY_400),
    )

    tabela_container = ft.Container(
        content=resultado,
        padding=10,
        expand=True,
    )

    scrollable_container = ft.Column(
        [tabela_container],
        scroll=ft.ScrollMode.AUTO,
        expand=True,
        height=400,  # Limit the height to prevent overflow
    )

    btn_salvar = ft.ElevatedButton(
        "Salvar como Excel",
        icon=ft.icons.SAVE,
        on_click=salvar_excel,
    )

    operacao_valor = ft.Dropdown(
        label="Operação para Valor",
        options=[
            ft.dropdown.Option("Soma"),
            ft.dropdown.Option("Contagem"),
            ft.dropdown.Option("Máximo"),
            ft.dropdown.Option("Mínimo"),
            ft.dropdown.Option("Média"),
        ],
        value="Soma",
        on_change=lambda _: atualizar_cubo(),
    )

    formatacao_valor = ft.Dropdown(
        label="Formatação do Valor",
        options=[
            ft.dropdown.Option("Número"),
            ft.dropdown.Option("Inteiro"),
            ft.dropdown.Option("Data"),
            ft.dropdown.Option("Hora"),
        ],
        value="Número",
        on_change=lambda _: atualizar_cubo(),
    )

    page.add(
        ft.Row([btn_selecionar_arquivos, btn_salvar]),
        ft.Column([ft.Text("Arquivos carregados:"), lista_arquivos]),
        ft.Row(
            [
                ft.Column([tipo_join, join_info_icon]),
                join_areas["left"],
                join_areas["right"],
            ]
        ),
        drags,
        ft.Row([drop_linha, drop_valor, drop_filtro]),
        ft.Row([operacao_valor, formatacao_valor, area_filtros]),
        scrollable_container,
    )

    atualizar_cubo()


ft.app(target=main)
