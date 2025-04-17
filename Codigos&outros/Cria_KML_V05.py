

# pip install pandas lxml openpyxl
# pip install geopy --trusted-host pypi.org --trusted-host files.pythonhosted.org

import os
import pandas as pd
from geopy.distance import geodesic
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import subprocess
import shutil

# Caminhos de entrada e saída
local_entrada = r"Z:\GEHITE\Principal\0-Gerencial\MAPAS\2025\KMZ_Estacoes"
local_saida = r"Z:\GEHITE\Principal\0-Gerencial\MAPAS\2025\KMZ_Estacoes"
# local_saida = r"G:/Drives compartilhados/GEHITE_Servidor-GO/0-Gerencial/MAPAS/2025/KMZ_Estacoes"
arquivo_entrada = "Coordenadas estações - SGIH_VERSAO-20250417.xlsx"

# Carregar o Excel
caminho_arquivo = os.path.join(local_entrada, arquivo_entrada)
dados = pd.read_excel(caminho_arquivo)

# Formatar data de criação
data_criacao = datetime.now().strftime("%Y.%m.%d")

# Função para calcular distância em metros


def dist_metros(coord1, coord2):
    return geodesic(coord1, coord2).meters

# Função para gerar KML

#########
### DADOS DE TESTE
#########

nome_roteiro=rot

#########
### FIM DADOS DE TESTE
#########

def gerar_kml(dados, nome_roteiro, local_saida, data_criacao):
    kml = ET.Element("kml", xmlns="http://www.opengis.net/kml/2.2")
    document = ET.SubElement(kml, "Document")
    ET.SubElement(document, "name").text = f"Roteiro {nome_roteiro}"
    ET.SubElement(document, "description").text = ""

    # Estilo com ícone circular laranja e rótulo visível
    style = ET.SubElement(document, "Style", id="estilo-completo")
    icon_style = ET.SubElement(style, "IconStyle")
    ET.SubElement(icon_style, "color").text = "ff00a5ff"  # laranja (AABBGGRR)
    ET.SubElement(icon_style, "scale").text = "1.1"
    icon = ET.SubElement(icon_style, "Icon")
    ET.SubElement(
        icon, "href").text = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"

    label_style = ET.SubElement(style, "LabelStyle")
    ET.SubElement(label_style, "scale").text = "1"

    folder = ET.SubElement(document, "Folder")
    ET.SubElement(folder, "name").text = f"Roteiro {nome_roteiro}"

    for _, row in dados.iterrows():
        estacao = row["Estacao"]
        rot = row["Rot"]

        has_P = not pd.isna(row["Latitude PLU"]) and not pd.isna(
            row["Longitude PLU"])
        has_F = not pd.isna(row["Latitude FLU"]) and not pd.isna(
            row["Longitude FLU"])

        usar_P = has_P and not has_F
        usar_F = has_F and not has_P
        usar_ambos = has_P and has_F

        if usar_ambos:
            dist = geodesic((row["Latitude PLU"], row["Longitude PLU"]),
                            (row["Latitude FLU"], row["Longitude FLU"])).meters
            if dist < 200:
                usar_P = False
                usar_F = True
                usar_ambos = False

        lat = row["Latitude PLU"] if usar_P else row["Latitude FLU"]
        lon = row["Longitude PLU"] if usar_P else row["Longitude FLU"]
        tipo = "PLU" if usar_P else "FLU" if usar_F else "PLU / FLU"
        codigo = row["CodigoPLU"] if usar_P else row[
            "CodigoFLU"] if usar_F else f"{row['CodigoPLU']} / {row['CodigoFLU']}"

        placemark = ET.SubElement(folder, "Placemark")
        ET.SubElement(placemark, "name").text = estacao
        ET.SubElement(placemark, "styleUrl").text = "#estilo-completo"

        descricao = f"""Código: {codigo}<br>
Tipo: {tipo}<br>
"""
        ET.SubElement(placemark, "description").text = descricao

        ext = ET.SubElement(placemark, "ExtendedData")
        for campo, valor in [
            ("Roteiro", rot), ("Tipo", tipo), ("Código", codigo),
            ("Latitude", lat), ("Longitude", lon)
        ]:
            data = ET.SubElement(ext, "Data", name=campo)
            ET.SubElement(data, "value").text = str(valor)

        point = ET.SubElement(placemark, "Point")
        ET.SubElement(point, "coordinates").text = f"{lon},{lat},0"

    kml_str = ET.tostring(kml, encoding="utf-8")
    kml_pretty = minidom.parseString(kml_str).toprettyxml(indent="  ")
    nome_arquivo = f"Roteiro_{nome_roteiro}_{data_criacao}.kml"
    caminho = os.path.join(local_saida, nome_arquivo)
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(kml_pretty)


rot=1   # TESTE

# Geração do KML dos rot
for rot in sorted(dados["Rot"].dropna().unique()):
    dados_rot = dados[dados["Rot"] == rot]
    gerar_kml(dados_rot, f"{int(rot):02d}", local_saida, data_criacao)



# Geração do KML completo
gerar_kml(dados, "Completo", local_saida, data_criacao)


################################################################################
#
# Cópia dos arquivos no GitHub
#
################################################################################

# Caminho para o diretório local onde estão os arquivos
os.chdir(local_saida)

# 🔹 Função para executar comandos shell


def run_command(cmd, check=True):
    try:
        result = subprocess.run(
            cmd, shell=True, check=check, text=True, capture_output=True)
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        print(f"Erro ao executar: {cmd}\n{e.stderr}")
        if check:
            raise
        return ""


# 🔹 Verificar se Git está instalado
versao_git = run_command("git --version", check=False)
if not versao_git or "not recognized" in versao_git or "127" in versao_git:
    raise EnvironmentError("❌ Git não está instalado ou não está no PATH.")
else:
    print("✔️ Git está instalado:", versao_git)

# 🔹 Inicializa Git se necessário
run_command("git init")

# 🔹 Define diretório como seguro
safe_path = os.path.normpath(local_saida).replace("\\", "/")
run_command(f'git config --global --add safe.directory "{safe_path}"')

# 🔹 Configura remoto, se necessário
remotes = run_command("git remote", check=False).splitlines()
if "origin" not in remotes:
    run_command(
        "git remote add origin https://github.com/maurocampostrindade/Roteiros_RHN-2025-kml.git")
    print("🔗 Remote 'origin' configurado.")
else:
    print("🔁 Remote 'origin' já configurado.")

# 🔹 Define branch como main
run_command("git branch -M main")

# 🔹 Remove todos os arquivos rastreados do repositório
run_command("git rm -r --cached .", check=False)

# 🔹 Commit da remoção
run_command(
    'git commit -m "🧹 Limpeza completa do repositório remoto antes da atualização"', check=False)

# 🔹 Adiciona os novos arquivos e faz novo commit
run_command("git add .")
run_command(
    'git commit -m "🚀 Atualização com dados SGIH-WEB"', check=False)

# 🔹 Push final (forçado para sobrescrever o repositório remoto)
run_command("git push origin main --force")
