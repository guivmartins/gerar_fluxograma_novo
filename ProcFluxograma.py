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
    """Detecta o encoding de um arquivo (usado como fallback)"""
    try:
        with open(filepath, 'rb') as f:
            resultado = chardet.detect(f.read())
            return resultado['encoding']
    except Exception:
        return 'utf-8'

def ler_excel_com_encoding(filepath):
    """
    Lê arquivo Excel com suporte a múltiplos formatos e tratamento de encoding.
    Suporta: .xlsx, .xls, .xlsm
    """
    file_extension = os.path.splitext(filepath)[1].lower()

    try:
        if file_extension in ['.xlsx', '.xlsm']:
            df = pd.read_excel(filepath, engine='openpyxl')
            return df

        elif file_extension == '.xls':
            try:
                df = pd.read_excel(filepath, engine='xlrd')
                return df
            except Exception as e:
                print(f"⚠️ Aviso: xlrd falhou, tentando openpyxl: {str(e)}")
                df = pd.read_excel(filepath, engine='openpyxl')
                return df

        else:
            raise ValueError(f"Formato não suportado: {file_extension}")

    except UnicodeDecodeError as e:
        print(f"⚠️ Erro de encoding detectado: {str(e)}")
        encoding_detectado = detectar_encoding(filepath)
        print(f"🔍 Encoding detectado: {encoding_detectado}")
        df = pd.read_excel(filepath, encoding=encoding_detectado)
        return df

    except Exception as e:
        raise ValueError(f"Erro ao ler arquivo Excel: {str(e)}. Verifique se o arquivo não está corrompido.")

def processar_para_drawflow(filepath):
    """
    Processar Excel e retornar dados para Drawflow com layout em CASCATA.
    CORRIGIDO: Evita sobreposição e garante círculos FIM
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
            raise ValueError(f"❌ Coluna obrigatória ausente no Excel: {col}")

    nome_processo = str(df["NOME PROCESSO"].dropna().unique()[0]).strip()

    nodes = {}
    connections = []
    node_id = 1

    # ESPAÇAMENTO AUMENTADO para evitar sobreposição
    x_spacing = 280  # Aumentado de 250 para 280
    y_spacing = 140  # Aumentado de 120 para 140

    # Rastrear posições Y usadas por coluna para evitar sobreposição
    coluna_y_max = {}  # {coluna: maior_y_usado}

    # Rastrear atividades
    atividade_info = {}  # {nome: {"coluna": X, "y": Y, "node_id": ID}}

    # PASSO 1: Criar nó de INÍCIO
    y_atual = 50
    inicio_key = "inicio"
    nodes[inicio_key] = {
        "id": node_id,
        "name": "Início",
        "type": "start",
        "pos_x": 50,
        "pos_y": y_atual
    }
    inicio_id = node_id
    node_id += 1
    coluna_y_max[0] = y_atual

    # PASSO 2: Identificar e criar atividade inicial
    df_inicio = df[df["ATIVIDADE INÍCIO"].str.strip().str.upper() == "SIM"]

    if not df_inicio.empty:
        atividade_inicial = str(df_inicio.iloc[0]["ATIVIDADE ORIGEM"]).strip()
        y_atual += y_spacing

        ativ_key = f"ativ_{atividade_inicial}"
        nodes[ativ_key] = {
            "id": node_id,
            "name": atividade_inicial,
            "type": "activity",
            "pos_x": 50,
            "pos_y": y_atual
        }

        atividade_info[atividade_inicial] = {
            "coluna": 0,
            "y": y_atual,
            "node_id": node_id
        }

        connections.append({
            "from": inicio_id,
            "to": node_id
        })

        coluna_y_max[0] = y_atual
        node_id += 1

    # PASSO 3: Processar cada linha
    for idx, row in df.iterrows():
        atividade_origem = str(row["ATIVIDADE ORIGEM"]).strip()
        procedimento = str(row["PROCEDIMENTO"]).strip()
        destino_raw = row["ATIVIDADE DESTINO"]
        atividade_destino = str(destino_raw).strip() if pd.notna(destino_raw) and str(destino_raw).strip() != "" else None

        # Garantir que atividade origem existe
        if atividade_origem not in atividade_info:
            # Criar em nova posição Y para evitar sobreposição
            coluna_nova = 0
            if coluna_nova not in coluna_y_max:
                coluna_y_max[coluna_nova] = 50

            y_nova = coluna_y_max[coluna_nova] + y_spacing

            ativ_key = f"ativ_{atividade_origem}"
            nodes[ativ_key] = {
                "id": node_id,
                "name": atividade_origem,
                "type": "activity",
                "pos_x": 50 + (coluna_nova * x_spacing),
                "pos_y": y_nova
            }

            atividade_info[atividade_origem] = {
                "coluna": coluna_nova,
                "y": y_nova,
                "node_id": node_id
            }

            coluna_y_max[coluna_nova] = y_nova
            node_id += 1

        # Info da atividade origem
        info_origem = atividade_info[atividade_origem]
        coluna_ativ = info_origem["coluna"]
        ativ_node_id = info_origem["node_id"]

        # PROCEDIMENTO (coluna seguinte)
        coluna_proc = coluna_ativ + 1
        x_proc = 50 + (coluna_proc * x_spacing)

        # Y do procedimento: usar próxima posição disponível na coluna
        if coluna_proc not in coluna_y_max:
            coluna_y_max[coluna_proc] = 50

        y_proc = coluna_y_max[coluna_proc] + y_spacing
        coluna_y_max[coluna_proc] = y_proc

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
            "from": ativ_node_id,
            "to": proc_node_id
        })

        # DESTINO (coluna seguinte ao procedimento)
        coluna_destino = coluna_proc + 1
        x_destino = 50 + (coluna_destino * x_spacing)

        if atividade_destino:
            # Verificar se é FIM - CORRIGIDO: verificação mais robusta
            if atividade_destino.upper() in ["FIM", "FINAL", "END"]:
                # Criar círculo FIM
                fim_key = f"fim_{idx}"
                nodes[fim_key] = {
                    "id": node_id,
                    "name": "Fim",
                    "type": "end",
                    "pos_x": x_destino,
                    "pos_y": y_proc
                }

                connections.append({
                    "from": proc_node_id,
                    "to": node_id
                })
                node_id += 1

            else:
                # Atividade destino normal
                destino_key = f"ativ_{atividade_destino}"

                if atividade_destino not in atividade_info:
                    # Y do destino: usar próxima posição disponível
                    if coluna_destino not in coluna_y_max:
                        coluna_y_max[coluna_destino] = 50

                    y_destino = coluna_y_max[coluna_destino] + y_spacing
                    coluna_y_max[coluna_destino] = y_destino

                    nodes[destino_key] = {
                        "id": node_id,
                        "name": atividade_destino,
                        "type": "activity",
                        "pos_x": x_destino,
                        "pos_y": y_destino
                    }

                    atividade_info[atividade_destino] = {
                        "coluna": coluna_destino,
                        "y": y_destino,
                        "node_id": node_id
                    }
                    node_id += 1

                # Conectar procedimento → atividade destino
                connections.append({
                    "from": proc_node_id,
                    "to": atividade_info[atividade_destino]["node_id"]
                })
        else:
            # Sem destino = criar FIM automaticamente
            fim_key = f"fim_auto_{idx}"
            nodes[fim_key] = {
                "id": node_id,
                "name": "Fim",
                "type": "end",
                "pos_x": x_destino,
                "pos_y": y_proc
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
    """Função original para gerar imagens estáticas com Graphviz"""
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
            raise ValueError(f"❌ Coluna obrigatória ausente no Excel: {col}")

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
