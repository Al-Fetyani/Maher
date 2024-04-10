import configparser
from pathlib import Path
from typing import List
from models import Template
import configparser

INI_FILE = Path("config.ini")


def create_ini_file():
    if not INI_FILE.exists():
        with open(INI_FILE, "w", encoding="utf-8") as file:
            file.write("")


def save_templates(templates: List[Template]):
    save_data({template.name: template.path for template in templates}, "templates")


def save_excel(file_path):
    file_name = Path(file_path).name

    save_data({file_name: file_path}, "excel")


def save_data(data, data_name):
    config = configparser.ConfigParser()
    config.read(INI_FILE, encoding="utf-8")
    config.remove_section(data_name)
    config[data_name] = data
    with open(INI_FILE, "w", encoding="utf-8") as file:
        config.write(file)


def load_templates() -> List[Template]:
    config = configparser.ConfigParser()
    config.read(INI_FILE, encoding="utf-8")
    if "templates" not in config:
        return []
    return [Template(name, path) for name, path in config["templates"].items()]


def load_excels():
    config = configparser.ConfigParser()
    config.read(INI_FILE, encoding="utf-8")
    if "excel" not in config:
        return None
    for _, path in config["excel"].items():
        if Path(path).exists():
            return path
