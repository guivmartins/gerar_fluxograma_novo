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
    # Detectar extensão do arquivo
    file_extension = os.path.splitext(filepath)[1].lower()

    try:
        # Tentar ler com pandas (suporta .xlsx e .xlsm nativamente)
        if file_extension in ['.xlsx', '.xlsm']:
            # openpyxl já lida com UTF-8 automaticamente
            df = pd.read_excel(filepath, engine='openpyxl')
            return df

        elif file_extension == '.xls':
            # Usar xlrd para arquivos XLS antigos (Excel 97-2003)
            try:
                df = pd.read_excel(filepath, engine='xlrd')
                return df
            except Exception as e:
                # Se xlrd falhar, tentar openpyxl como fallback
                print(f"⚠️ Aviso: xlrd falhou, tentando openpyxl: {str(e)}")
                df = pd.read_excel(filepath, engine='openpyxl')
                return df

        else:
            raise ValueError(f"Formato não suportado: {file_extension}")

    except UnicodeDecodeError as e:
        # Tratamento de erro de encoding (raramente acontece com Excel)
        print(f"⚠️ Erro de encoding detectado: {str(e)}")
        encoding_detectado = detectar_encoding(filepath)
        print(f"🔍 Encoding detectado: {encoding_detectado}")

        # Tentar novamente com encoding detectado
        df = pd.read_excel(filepath, encoding=encoding_detectado)
        return df

    except Exception as e:
        # Erro genérico com mensagem amigável
        raise ValueError(f"Erro ao ler arquivo Excel: {str(e)}. Verifique se o arquivo não está corrompido.")

def processar_para_drawflow(filepath):
    """Processar Excel e retornar dados para Drawflow"""
    # Ler arquivo com tratamento de encoding
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
    pos_x = 50
    pos_y = 50
    y_spacing = 120
    x_spacing = 250

    df_agrupado = df.groupby(["ATIVIDADE ORIGEM", "PROCEDIMENTO"]).agg({
        "ATIVIDADE INÍCIO": "first",
        "ATIVIDADE DESTINO": lambda x: [str(i) for i in x.dropna()]
    }).reset_index()

    inicio_criado = False

    # Criar nós
    for idx, row in df_agrupado.iterrows():
        atividade = str(row["ATIVIDADE ORIGEM"]).strip()
        procedimento = str(row["PROCEDIMENTO"]).strip()
        destinos = row["ATIVIDADE DESTINO"]
        inicio = str(row["ATIVIDADE INÍCIO"]).strip().upper() if pd.notna(row["ATIVIDADE INÍCIO"]) else "NAO"

        current_y = pos_y + (idx * y_spacing)

        # Nó de início
        if inicio == "SIM" and not inicio_criado:
            inicio_key = "inicio"
            nodes[inicio_key] = {
                "id": node_id,
                "name": "Início",
                "type": "start",
                "pos_x": pos_x,
                "pos_y": current_y
            }
            node_id += 1
            inicio_criado = True

        # Nó de atividade
        atividade_key = f"ativ_{atividade}"
        if atividade_key not in nodes:
            nodes[atividade_key] = {
                "id": node_id,
                "name": atividade,
                "type": "activity",
                "pos_x": pos_x + x_spacing,
                "pos_y": current_y
            }

            # Conectar início à primeira atividade se for o caso
            if inicio == "SIM" and "inicio" in nodes:
                connections.append({
                    "from": nodes["inicio"]["id"],
                    "to": node_id
                })

            node_id += 1

        # Nó de procedimento
        proc_key = f"proc_{atividade}_{procedimento}"
        if proc_key not in nodes:
            nodes[proc_key] = {
                "id": node_id,
                "name": procedimento,
                "type": "procedure",
                "pos_x": pos_x + (x_spacing * 2),
                "pos_y": current_y
            }
            node_id += 1

        # Conexão: atividade -> procedimento
        connections.append({
            "from": nodes[atividade_key]["id"],
            "to": nodes[proc_key]["id"]
        })

        # Nós de destino e conexões
        if destinos and len(destinos) > 0 and destinos[0]:
            for i, destino in enumerate(destinos):
                destino_key = f"ativ_{destino}"
                if destino_key not in nodes:
                    nodes[destino_key] = {
                        "id": node_id,
                        "name": destino,
                        "type": "activity",
                        "pos_x": pos_x + (x_spacing * 3),
                        "pos_y": current_y + (i * 80)
                    }
                    node_id += 1

                connections.append({
                    "from": nodes[proc_key]["id"],
                    "to": nodes[destino_key]["id"]
                })
        else:
            # Nó de fim
            fim_key = f"fim_{idx}"
            if fim_key not in nodes:
                nodes[fim_key] = {
                    "id": node_id,
                    "name": "Fim",
                    "type": "end",
                    "pos_x": pos_x + (x_spacing * 3),
                    "pos_y": current_y
                }
                node_id += 1

            connections.append({
                "from": nodes[proc_key]["id"],
                "to": nodes[fim_key]["id"]
            })

    return {
        "nome_processo": nome_processo,
        "nodes": list(nodes.values()),
        "connections": connections
    }

def gerar_fluxograma(filepath):
    """Função original para gerar imagens estáticas com Graphviz"""
    # Usar a nova função de leitura com encoding
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
