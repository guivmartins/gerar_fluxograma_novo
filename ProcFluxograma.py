# -*- coding: utf-8 -*-
import pandas as pd
from graphviz import Digraph
import textwrap
import os
import chardet

# Cores
COLOR_EDGE = "#00796B"
COLOR_ACTIVITY_FILL = "#00AE9D"
COLOR_ACTIVITY_BORDER = "#006B5A"
COLOR_ACTIVITY_FONT = "white"
COLOR_PROC_FILL = "#E8E8E8"
COLOR_PROC_FONT = "#666666"
COLOR_START_FILL = "#87C2BC"
COLOR_START_BORDER = "#00A896"
COLOR_START_FONT = "#006B5A"
COLOR_END_FILL = "#87C2BC"
COLOR_END_BORDER = "#00A896"
COLOR_END_FONT = "#006B5A"

def wrap_label(text, max_len=15):
    if text is None:
        return ""
    text = str(text).strip()
    if not text:
        return ""
    return "\n".join(textwrap.wrap(text, max_len))

def detectar_encoding(filepath):
    """Detecta o encoding de um arquivo"""
    try:
        with open(filepath, 'rb') as f:
            resultado = chardet.detect(f.read())
            return resultado['encoding']
    except Exception:
        return 'utf-8'

def ler_excel_com_encoding(filepath):
    """Lê arquivo Excel com suporte a múltiplos formatos"""
    file_extension = os.path.splitext(filepath)[1].lower()

    try:
        if file_extension in ['.xlsx', '.xlsm']:
            df = pd.read_excel(filepath, engine='openpyxl')
            return df
        elif file_extension == '.xls':
            try:
                df = pd.read_excel(filepath, engine='xlrd')
                return df
            except Exception:
                df = pd.read_excel(filepath, engine='openpyxl')
                return df
        else:
            raise ValueError(f"Formato não suportado: {file_extension}")
    except UnicodeDecodeError:
        encoding_detectado = detectar_encoding(filepath)
        df = pd.read_excel(filepath, encoding=encoding_detectado)
        return df
    except Exception as e:
        raise ValueError(f"Erro ao ler arquivo Excel: {str(e)}")

def processar_para_drawflow(filepath):
    """
    Layout em CASCATA - SEM SOBREPOSIÇÃO

    Col 1: Início
    Col 2: Atividades
    Col 3: Procedimentos
    Col 4: Atividades + Fins (SEM SOBREPOSIÇÃO)
    Col 5: Procedimentos
    Col 6: Atividades + Fins (SEM SOBREPOSIÇÃO)
    ...
    """
    df = ler_excel_com_encoding(filepath)

    colunas_necessarias = [
        "NOME PROCESSO",
        "ATIVIDADE INÍCIO",
        "ATIVIDADE ORIGEM",
        "PROCEDIMENTO",
        "ATIVIDADE DESTINO",
    ]

    for col in colunas_necessarias:
        if col not in df.columns:
            raise ValueError(f"❌ Coluna obrigatória ausente: {col}")

    nome_processo = str(df["NOME PROCESSO"].dropna().unique()[0]).strip()

    nodes = {}
    connections = []
    node_id = 1

    # Espaçamento AUMENTADO para evitar sobreposição
    x_spacing = 300  # Espaço horizontal entre colunas
    y_spacing = 150  # Espaço vertical entre nós

    # Rastrear Y atual de cada coluna (CRÍTICO para evitar sobreposição)
    coluna_y_atual = {}

    def get_next_y(coluna):
        """
        Retorna próximo Y disponível para uma coluna.
        GARANTE que nunca haverá sobreposição.
        """
        if coluna not in coluna_y_atual:
            coluna_y_atual[coluna] = 50
        else:
            coluna_y_atual[coluna] += y_spacing
        return coluna_y_atual[coluna]

    # Rastrear atividades
    atividade_info = {}  # {nome: {"coluna": X, "node_id": ID}}

    # === COLUNA 1: INÍCIO ===
    y_inicio = get_next_y(1)
    inicio_key = "inicio"
    nodes[inicio_key] = {
        "id": node_id,
        "name": "Início",
        "type": "start",
        "pos_x": 50,
        "pos_y": y_inicio
    }
    inicio_id = node_id
    node_id += 1

    # === COLUNA 2: ATIVIDADE INICIAL ===
    df_inicio = df[df["ATIVIDADE INÍCIO"].str.strip().str.upper() == "SIM"]

    if not df_inicio.empty:
        atividade_inicial = str(df_inicio.iloc[0]["ATIVIDADE ORIGEM"]).strip()

        y_ativ = get_next_y(2)
        x_ativ = 50 + (1 * x_spacing)

        ativ_key = f"ativ_{atividade_inicial}"
        nodes[ativ_key] = {
            "id": node_id,
            "name": atividade_inicial,
            "type": "activity",
            "pos_x": x_ativ,
            "pos_y": y_ativ
        }

        atividade_info[atividade_inicial] = {
            "coluna": 2,
            "node_id": node_id
        }

        connections.append({
            "from": inicio_id,
            "to": node_id
        })
        node_id += 1

    # === PROCESSAR TODAS AS LINHAS ===
    for idx, row in df.iterrows():
        atividade_origem = str(row["ATIVIDADE ORIGEM"]).strip()
        procedimento = str(row["PROCEDIMENTO"]).strip()
        destino_raw = row["ATIVIDADE DESTINO"]
        atividade_destino = str(destino_raw).strip() if pd.notna(destino_raw) and str(destino_raw).strip() else None

        # Garantir que atividade origem existe
        if atividade_origem not in atividade_info:
            # Criar atividade em coluna PAR (2, 4, 6...)
            # Encontrar próxima coluna par disponível
            colunas_atividades = [info["coluna"] for info in atividade_info.values()]
            proxima_coluna = 2
            while proxima_coluna in colunas_atividades:
                proxima_coluna += 2

            y_ativ = get_next_y(proxima_coluna)
            x_ativ = 50 + ((proxima_coluna - 1) * x_spacing)

            ativ_key = f"ativ_{atividade_origem}"
            nodes[ativ_key] = {
                "id": node_id,
                "name": atividade_origem,
                "type": "activity",
                "pos_x": x_ativ,
                "pos_y": y_ativ
            }

            atividade_info[atividade_origem] = {
                "coluna": proxima_coluna,
                "node_id": node_id
            }
            node_id += 1

        # Info da atividade origem
        coluna_atividade = atividade_info[atividade_origem]["coluna"]
        atividade_node_id = atividade_info[atividade_origem]["node_id"]

        # === COLUNA ÍMPAR: PROCEDIMENTO ===
        coluna_proc = coluna_atividade + 1  # 3, 5, 7, 9...
        y_proc = get_next_y(coluna_proc)  # GARANTE Y único
        x_proc = 50 + ((coluna_proc - 1) * x_spacing)

        proc_key = f"proc_{idx}"
        nodes[proc_key] = {
            "id": node_id,
            "name": procedimento,
            "type": "procedure",
            "pos_x": x_proc,
            "pos_y": y_proc
        }
        proc_node_id = node_id
        node_id += 1

        # Conectar atividade → procedimento
        connections.append({
            "from": atividade_node_id,
            "to": proc_node_id
        })

        # === COLUNA PAR: ATIVIDADE DESTINO ou FIM ===
        coluna_destino = coluna_proc + 1  # 4, 6, 8, 10...
        x_destino = 50 + ((coluna_destino - 1) * x_spacing)

        if atividade_destino:
            # Verificar se é FIM
            if atividade_destino.upper() in ["FIM", "FINAL", "END"]:
                # === CÍRCULO FIM ===
                y_fim = get_next_y(coluna_destino)  # Y ÚNICO - SEM SOBREPOSIÇÃO

                fim_key = f"fim_{idx}"
                nodes[fim_key] = {
                    "id": node_id,
                    "name": "Fim",
                    "type": "end",
                    "pos_x": x_destino,
                    "pos_y": y_fim
                }

                connections.append({
                    "from": proc_node_id,
                    "to": node_id
                })
                node_id += 1
            else:
                # === ATIVIDADE DESTINO ===
                if atividade_destino not in atividade_info:
                    # Criar nova atividade
                    y_dest = get_next_y(coluna_destino)  # Y ÚNICO - SEM SOBREPOSIÇÃO

                    destino_key = f"ativ_{atividade_destino}"
                    nodes[destino_key] = {
                        "id": node_id,
                        "name": atividade_destino,
                        "type": "activity",
                        "pos_x": x_destino,
                        "pos_y": y_dest
                    }

                    atividade_info[atividade_destino] = {
                        "coluna": coluna_destino,
                        "node_id": node_id
                    }
                    node_id += 1

                # Conectar procedimento → atividade destino
                connections.append({
                    "from": proc_node_id,
                    "to": atividade_info[atividade_destino]["node_id"]
                })
        else:
            # === SEM DESTINO = FIM AUTOMÁTICO ===
            y_fim = get_next_y(coluna_destino)  # Y ÚNICO - SEM SOBREPOSIÇÃO

            fim_key = f"fim_auto_{idx}"
            nodes[fim_key] = {
                "id": node_id,
                "name": "Fim",
                "type": "end",
                "pos_x": x_destino,
                "pos_y": y_fim
            }

            connections.append({
                "from": proc_node_id,
                "to": node_id
            })
            node_id += 1

    return {
        "nome_processo": nome_processo,
        "nodes": list(nodes.values()),
        "connections": connections
    }

def gerar_fluxograma(filepath):
    """Função para gerar imagens estáticas com Graphviz"""
    df = ler_excel_com_encoding(filepath)

    colunas_necessarias = [
        "NOME PROCESSO",
        "ATIVIDADE INÍCIO",
        "ATIVIDADE ORIGEM",
        "PROCEDIMENTO",
        "ATIVIDADE DESTINO",
    ]

    for col in colunas_necessarias:
        if col not in df.columns:
            raise ValueError(f"❌ Coluna obrigatória ausente: {col}")

    nome_processo = str(df["NOME PROCESSO"].dropna().unique()[0]).strip()

    dot = Digraph(comment="Fluxograma", format="png")
    dot.attr(rankdir="LR")
    dot.attr(label=nome_processo, labelloc="t", fontsize="20",
             fontname="Roboto", fontcolor="black")
    dot.attr(splines="ortho")
    dot.attr("edge", color=COLOR_EDGE, penwidth="1.5", arrowsize="0.8")

    dot.node_attr.update(fontname="Roboto")
    dot.edge_attr.update(fontname="Roboto")

    nos_criados = set()
    arestas_criadas = set()

    df_agrupado = df.groupby(["ATIVIDADE ORIGEM", "PROCEDIMENTO"]).agg({
        "ATIVIDADE INÍCIO": "first",
        "ATIVIDADE DESTINO": lambda x: [str(i) for i in x.dropna()]
    }).reset_index()

    for _, row in df_agrupado.iterrows():
        raw_atividade = row["ATIVIDADE ORIGEM"]
        raw_procedimento = row["PROCEDIMENTO"]
        raw_destinos = row["ATIVIDADE DESTINO"]
        inicio_flag = str(row["ATIVIDADE INÍCIO"]).strip().upper() if pd.notna(row["ATIVIDADE INÍCIO"]) else "NAO"

        atividade = wrap_label(raw_atividade)
        procedimento = wrap_label(raw_procedimento)

        def safe_id(text, prefix="n"):
            import re
            t = (prefix + "_" + (text if text else "vazio")).replace(" ", "_")
            t = re.sub(r'[^0-9A-Za-z_áàãâéêíóôõúçÁÀÃÂÉÊÍÓÔÕÚÇ\-]', '', t)
            return t

        atividade_id = safe_id(str(raw_atividade), prefix="act")
        proc_id = safe_id(f"{raw_atividade}__{raw_procedimento}", prefix="proc")

        if inicio_flag == "SIM":
            inicio_node = f"inicio_{atividade_id}"
            if inicio_node not in nos_criados:
                dot.node(inicio_node, "Início", shape="rect", style="rounded,filled",
                         fillcolor=COLOR_START_FILL, color=COLOR_START_BORDER,
                         fontname="Roboto", fontcolor=COLOR_START_FONT, fontsize="12",
                         width="0.9", height="0.5", fixedsize="true")
                nos_criados.add(inicio_node)
            if (inicio_node, atividade_id) not in arestas_criadas:
                dot.edge(inicio_node, atividade_id)
                arestas_criadas.add((inicio_node, atividade_id))

        if atividade_id not in nos_criados:
            dot.node(atividade_id, atividade if atividade else str(raw_atividade),
                     shape="rect", style="rounded,filled", fillcolor=COLOR_ACTIVITY_FILL,
                     color=COLOR_ACTIVITY_BORDER, fontname="Roboto", fontcolor=COLOR_ACTIVITY_FONT,
                     fontsize="12", width="1.5", height="0.9", fixedsize="true")
            nos_criados.add(atividade_id)

        if proc_id not in nos_criados:
            dot.node(proc_id, procedimento if procedimento else str(raw_procedimento),
                     shape="rect", style="filled", fillcolor=COLOR_PROC_FILL, color="#d0d0d0",
                     fontname="Roboto", fontcolor=COLOR_PROC_FONT, fontsize="8",
                     width="1.2", height="0.6", fixedsize="true")
            nos_criados.add(proc_id)

        if (atividade_id, proc_id) not in arestas_criadas:
            dot.edge(atividade_id, proc_id)
            arestas_criadas.add((atividade_id, proc_id))

        if raw_destinos:
            for destino_text in raw_destinos:
                destino = wrap_label(destino_text)
                destino_id = safe_id(destino_text, prefix="act")
                if destino_id not in nos_criados:
                    dot.node(destino_id, destino, shape="rect", style="rounded,filled",
                             fillcolor=COLOR_ACTIVITY_FILL, color=COLOR_ACTIVITY_BORDER,
                             fontname="Roboto", fontcolor=COLOR_ACTIVITY_FONT,
                             fontsize="12", width="1.5", height="0.9", fixedsize="true")
                    nos_criados.add(destino_id)
                if (proc_id, destino_id) not in arestas_criadas:
                    dot.edge(proc_id, destino_id)
                    arestas_criadas.add((proc_id, destino_id))
        else:
            fim_node = f"fim_{atividade_id}"
            if fim_node not in nos_criados:
                dot.node(fim_node, "Fim", shape="rect", style="rounded,filled",
                         fillcolor=COLOR_END_FILL, color=COLOR_END_BORDER,
                         fontname="Roboto", fontcolor=COLOR_END_FONT,
                         fontsize="12", width="0.9", height="0.5", fixedsize="true")
                nos_criados.add(fim_node)
            if (proc_id, fim_node) not in arestas_criadas:
                dot.edge(proc_id, fim_node)
                arestas_criadas.add((proc_id, fim_node))

    output_dir = "static"
    os.makedirs(output_dir, exist_ok=True)
    dot.render(os.path.join(output_dir, "fluxograma"), format="png", cleanup=True)
    dot.render(os.path.join(output_dir, "fluxograma"), format="pdf", cleanup=True)
    dot.render(os.path.join(output_dir, "fluxograma"), format="svg", cleanup=True)
