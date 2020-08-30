import configparser
import os


def path_to_settings_file():
    """
    :return: возвращает путь к файлу настроек settings.ini
    """
    return os.path.abspath("" + "settings.ini")


def create_settings_file(path=path_to_settings_file()):
    """
    Создает файл настроек settings.ini по умолчанию, если он отсутствует
    """
    if not os.path.exists(path):
        config = configparser.ConfigParser()
        config.add_section("Settings")
        config.set("Settings", "local_acts_path", "C:\\")
        config.set("Settings", "general_acts_path", "C:\\")
        config.set("Settings", "path_to_docx_tmpl", "C:\\")

        with open(path, "w") as config_file:
            config.write(config_file)


def get_local_acts_path_folder(path=path_to_settings_file()):
    """
    :param path: путь до файла settings.ini, передается по умолчанию
    :return: возвращает путь до папки с локальными актами из файла settings.ini
    """
    config = configparser.ConfigParser()
    config.read(path)
    if os.path.exists(config.get("Settings", "local_acts_path")):
        return config.get("Settings", "local_acts_path")
    return "C:\\"


def get_general_acts_path_folder(path=path_to_settings_file()):
    """
    :param path: путь до файла settings.ini, передается по умолчанию
    :return: возвращает путь до папки с общими актами из файла settings.ini
    """
    config = configparser.ConfigParser()
    config.read(path)
    return config.get("Settings", "general_acts_path")


def get_docx_templ_path(path=path_to_settings_file()):
    """
    :param path: путь до файла settings.ini, передается по умолчанию
    :return: возвращает путь файла шаблона docx из файла settings.ini
    """
    config = configparser.ConfigParser()
    config.read(path)
    return config.get("Settings", "path_to_docx_tmpl")


def set_local_acts_path_folder_in_settings(str_path, path=path_to_settings_file()):
    """
    записывает путь к папке локальных актов в файл настроек settings.ini
    :param str_path: путь в виде строки
    :param path: путь до файла settings.ini передается по умолчанию
    """
    config = configparser.ConfigParser()
    config.read(path)
    config.set("Settings", "local_acts_path", str_path)
    with open(path, "w") as config_file:
        config.write(config_file)


def set_general_acts_path_folder_in_settings(str_path, path=path_to_settings_file()):
    """
    записывает путь к папке общих актов в файл настроек settings.ini
    :param str_path: путь в виде строки
    :param path: путь до файла settings.ini передается по умолчанию
    """
    config = configparser.ConfigParser()
    config.read(path)
    config.set("Settings", "general_acts_path", str_path)
    with open(path, "w") as config_file:
        config.write(config_file)


def set_path_to_docx_templ_in_settings(str_path, path=path_to_settings_file()):
    """
    записывает путь к файлу шаблона docx в файл настроек settings.ini
    :param str_path: путь в виде строки
    :param path: путь до файла settings.ini передается по умолчанию
    """
    config = configparser.ConfigParser()
    config.read(path)
    config.set("Settings", "path_to_docx_tmpl", str_path)
    with open(path, "w") as config_file:
        config.write(config_file)
