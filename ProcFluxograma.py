# -*- coding: utf-8 -*-
import pandas as pd
import os
import chardet

def detectar_encoding(filepath):
    try:
        with open(filepath, 'rb') as f:
            resultado = chardet.detect(f.read())
            return resultado['encoding']
    except Exception:
        return 'utf-8'

def ler_excel_com_encoding(filepath):
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
    df = ler_excel_com_encoding(filepath)
    colunas_necessarias = [
        "NOME PROCESSO", "ATIVIDADE INÍCIO",
        "ATIVIDADE ORIGEM", "PROCEDIMENTO", "ATIVIDADE DESTINO"
    ]
    for col in colunas_necessarias:
        if col not in df.columns:
            raise ValueError(f"❌ Coluna obrigatória ausente: {col}")

    nome_processo = str(df["NOME PROCESSO"].dropna().unique()[0]).strip()
    nodes = {}
    connections = []
    node_id = 1
    x_spacing = 300
    y_spacing = 150
    coluna_y_atual = {}
    atividade_info = {}

    def get_next_y(coluna):
        if coluna not in coluna_y_atual:
            coluna_y_atual[coluna] = 50
        else:
            coluna_y_atual[coluna] += y_spacing
        return coluna_y_atual[coluna]

    # Início
    y_inicio = get_next_y(1)
    nodes["inicio"] = {
        "id": node_id, "name": "Início", "type": "start", "pos_x": 50, "pos_y": y_inicio
    }
    inicio_id = node_id
    node_id += 1

    # Atividade inicial
    df_inicio = df[df["ATIVIDADE INÍCIO"].str.strip().str.upper() == "SIM"]
    if not df_inicio.empty:
        atividade_inicial = str(df_inicio.iloc[0]["ATIVIDADE ORIGEM"]).strip()
        y_ativ = get_next_y(2)
        x_ativ = 50 + (1 * x_spacing)
        ativ_key = f"ativ_{atividade_inicial}"
        nodes[ativ_key] = {
            "id": node_id, "name": atividade_inicial, "type": "activity",
            "pos_x": x_ativ, "pos_y": y_ativ
        }
        atividade_info[atividade_inicial] = {
            "coluna": 2, "node_id": node_id
        }
        connections.append({
            "from": inicio_id, "to": node_id
        })
        node_id += 1

    for idx, row in df.iterrows():
        atividade_origem = str(row["ATIVIDADE ORIGEM"]).strip()
        procedimento = str(row["PROCEDIMENTO"]).strip()
        destino_raw = row["ATIVIDADE DESTINO"]
        atividade_destino = str(destino_raw).strip() if pd.notna(destino_raw) and str(destino_raw).strip() else None

        if atividade_origem not in atividade_info:
            colunas_atividades = [info["coluna"] for info in atividade_info.values()]
            proxima_coluna = 2
            while proxima_coluna in colunas_atividades:
                proxima_coluna += 2
            y_ativ = get_next_y(proxima_coluna)
            x_ativ = 50 + ((proxima_coluna - 1) * x_spacing)
            ativ_key = f"ativ_{atividade_origem}"
            nodes[ativ_key] = {
                "id": node_id, "name": atividade_origem, "type": "activity",
                "pos_x": x_ativ, "pos_y": y_ativ
            }
            atividade_info[atividade_origem] = {
                "coluna": proxima_coluna, "node_id": node_id
            }
            node_id += 1

        coluna_atividade = atividade_info[atividade_origem]["coluna"]
        atividade_node_id = atividade_info[atividade_origem]["node_id"]

        coluna_proc = coluna_atividade + 1
        y_proc = get_next_y(coluna_proc)
        x_proc = 50 + ((coluna_proc - 1) * x_spacing)
        proc_key = f"proc_{idx}"
        nodes[proc_key] = {
            "id": node_id, "name": procedimento, "type": "procedure",
            "pos_x": x_proc, "pos_y": y_proc
        }
        proc_node_id = node_id
        node_id += 1

        connections.append({
            "from": atividade_node_id, "to": proc_node_id
        })

        coluna_destino = coluna_proc + 1
        x_destino = 50 + ((coluna_destino - 1) * x_spacing)
        if atividade_destino:
            if atividade_destino.upper() in ["FIM", "FINAL", "END"]:
                y_fim = get_next_y(coluna_destino)
                fim_key = f"fim_{idx}"
                nodes[fim_key] = {
                    "id": node_id, "name": "Fim", "type": "end",
                    "pos_x": x_destino, "pos_y": y_fim
                }
                connections.append({
                    "from": proc_node_id, "to": node_id
                })
                node_id += 1
            else:
                if atividade_destino not in atividade_info:
                    y_dest = get_next_y(coluna_destino)
                    nodes[f"ativ_{atividade_destino}"] = {
                        "id": node_id, "name": atividade_destino, "type": "activity",
                        "pos_x": x_destino, "pos_y": y_dest
                    }
                    atividade_info[atividade_destino] = {
                        "coluna": coluna_destino, "node_id": node_id
                    }
                    node_id += 1
                destino_node_id = atividade_info[atividade_destino]["node_id"]
                connections.append({
                    "from": proc_node_id, "to": destino_node_id
                })
        else:
            y_fim = get_next_y(coluna_destino)
            fim_key = f"fim_auto_{idx}"
            nodes[fim_key] = {
                "id": node_id, "name": "Fim", "type": "end",
                "pos_x": x_destino, "pos_y": y_fim
            }
            connections.append({
                "from": proc_node_id, "to": node_id
            })
            node_id += 1

    return {
        "nome_processo": nome_processo,
        "nodes": list(nodes.values()),
        "connections": connections
    }
