# -*- coding: utf-8 -*-
"""
Script Orquestrador Mestre: Build do Dashboard de Liquidação BCMS
Gera a aplicação autocontida HTML5 para GitHub Pages com dados do SIAFI/PCASP injetados inline.
"""

import os
import sys
import json
import pathlib
import subprocess

REPO_DIR = pathlib.Path(__file__).resolve().parent
SCRIPTS_DIR = REPO_DIR / "scripts"
DATA_JSON = REPO_DIR / "data" / "liquidacao_data.json"
HTML_FILE = REPO_DIR / "site" / "index.html"

def build():
    print("==================================================================")
    print("  BUILD DASHBOARD LIQUIDAÇÃO & RESTOS A PAGAR · BCMS (2026)")
    print("  Base de Apoio Logístico (Ba Ap Log) · SIAFI / Tesouro Gerencial")
    print("==================================================================")

    # 1. Executa o extrator de dados
    extrator_py = SCRIPTS_DIR / "extrair_liquidacao_bcms.py"
    print(f"1. Executando extrator: {extrator_py.name}...")
    cmd = [sys.executable, str(extrator_py)]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"Erro ao extrair dados:\n{res.stderr}")
        sys.exit(1)
    print(res.stdout.strip())

    # 2. Carrega os dados gerados
    print("2. Lendo base estruturada JSON...")
    if not DATA_JSON.exists():
        print("Erro: liquidacao_data.json não foi gerado.")
        sys.exit(1)
    data_content = DATA_JSON.read_text(encoding="utf-8")

    # 3. Injeta no index.html
    print("3. Injetando dados inline no site/index.html...")
    if not HTML_FILE.exists():
        print("Erro: site/index.html não encontrado.")
        sys.exit(1)
    
    html_content = HTML_FILE.read_text(encoding="utf-8")
    
    # Substitui a declaração da variável LIQUIDACAO_DATA
    target_var = "var LIQUIDACAO_DATA = null;"
    if target_var in html_content:
        replacement = f"var LIQUIDACAO_DATA = {data_content};"
        new_html = html_content.replace(target_var, replacement, 1)
        HTML_FILE.write_text(new_html, encoding="utf-8")
        print("-> Dados injetados com sucesso!")
    else:
        # Se já foi injetado antes, substitui via regex
        import re
        new_html = re.sub(
            r'var LIQUIDACAO_DATA\s*=\s*\{.*?\};',
            f'var LIQUIDACAO_DATA = {data_content};',
            html_content,
            flags=re.DOTALL
        )
        HTML_FILE.write_text(new_html, encoding="utf-8")
        print("-> Dados reinjetados com sucesso via regex!")

    print("==================================================================")
    print("  BUILD CONCLUÍDO COM SUCESSO!")
    print(f"  Arquivo pronto para GitHub Pages: {HTML_FILE}")
    print("==================================================================")

if __name__ == "__main__":
    build()
