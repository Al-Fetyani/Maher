import configparser
from pathlib import Path
from typing import List, Union
from models import Template
import configparser

INI_FILE = Path(__file__).parent / "config.ini"


def save_templates(templates: List[Template]):
    save_data({template.name: template.path for template in templates}, "templates")


def save_data(data, data_name):
    config = configparser.ConfigParser()
    config.read(INI_FILE, encoding="utf-8")
    config.remove_section(data_name)  # Clear existing data
    config[data_name] = data
    with open(INI_FILE, "w", encoding="utf-8") as file:  # Use "w" mode to overwrite the file
        config.write(file)


def load_templates() -> List[Template]:
    if not INI_FILE.exists():
        return []
    config = configparser.ConfigParser()
    config.read(INI_FILE, encoding="utf-8")
    return [Template(name, path) for name, path in config["templates"].items()]
