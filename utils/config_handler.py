import json
import os

def get_caminho_config():
    # Cria uma pasta oculta chamada ".crea_bot" na pasta do usuário (cross-platform)
    pasta_config = os.path.join(os.path.expanduser("~"), ".crea_bot")
    os.makedirs(pasta_config, exist_ok=True)
    return os.path.join(pasta_config, "config.json")

def salvar_config(config):
    try:
        caminho = get_caminho_config()
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4)
    except Exception as e:
        print(f"Erro ao salvar config: {e}")

def carregar_config():
    try:
        caminho = get_caminho_config()
        if os.path.exists(caminho):
            with open(caminho, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        print(f"Erro ao carregar config: {e}")
    return {}
