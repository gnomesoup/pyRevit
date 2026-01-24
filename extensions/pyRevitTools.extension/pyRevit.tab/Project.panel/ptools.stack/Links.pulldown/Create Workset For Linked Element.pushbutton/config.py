# -*- coding: UTF-8 -*-
from pyrevit import script, forms
from pyrevit.userconfig import user_config
from script import main
my_config = script.get_config()

set_type_ws = my_config.get_option("set_type_ws", False)
set_all = my_config.get_option("set_all", False)
custom_prefix_for_rvt = my_config.get_option("custom_prefix_for_rvt", False)
custom_prefix_for_dwg = my_config.get_option("custom_prefix_for_dwg", False)


class MyOption(forms.TemplateListItem):
    @property
    def name(self):
        return str(self.item)


pyrevit_locale = user_config.user_locale  # type: str
# Translation dictionaries
translations = {
    "en_us": {
        "set_type_ws": "Set Workset for Type",
        "set_all": "Collect all Links",
        "custom_prefix_for_rvt": "Custom Prefix for RVT",
        "custom_prefix_for_dwg": "Custom Prefix for DWG",
        "results_title": "Select Options",
        "results_button_name": "Save Selection",
        "custom_prefix_rvt_value": "Pick a Prefix for RVTs",
        "custom_prefix_dwg_value": "Pick a Prefix for DWGs",
    },
    "fr_fr": {
        "set_type_ws": "Définir le sous-projet pour le type",
        "set_all": "Collecter tous les liens",
        "custom_prefix_for_rvt": "Préfixe personnalisé pour RVT",
        "custom_prefix_for_dwg": "Préfixe personnalisé pour DWG",
        "results_title": "Sélectionner les options",
        "results_button_name": "Enregistrer la sélection",
        "custom_prefix_rvt_value": "Choisissez un préfixe pour les RVT",
        "custom_prefix_dwg_value": "Choisissez un préfixe pour les DWG",
    },
    "ru": {
        "set_type_ws": "Назначить рабочий набор для типа",
        "set_all": "Собрать все связи",
        "custom_prefix_for_rvt": "Пользовательский префикс для RVT-связей",
        "custom_prefix_for_dwg": "Пользовательский префикс для DWG-связей",
        "results_title": "Выберите параметры",
        "results_button_name": "Сохранить выбор",
        "custom_prefix_rvt_value": "Выберите префикс для RVT-связей",
        "custom_prefix_dwg_value": "Выберите префикс для DWG-связей",
    },
    "chinese_s": {
        "set_type_ws": "为类型设置工作集",
        "set_all": "收集所有链接",
        "custom_prefix_for_rvt": "RVT 的自定义前缀",
        "custom_prefix_for_dwg": "DWG 的自定义前缀",
        "results_title": "选择选项",
        "results_button_name": "保存选择",
        "custom_prefix_rvt_value": "为 RVT 选择一个前缀",
        "custom_prefix_dwg_value": "为 DWG 选择一个前缀",
    },
    "es_es": {
        "set_type_ws": "Establecer subproyecto para tipo",
        "set_all": "Recopilar todos los vínculos",
        "custom_prefix_for_rvt": "Prefijo personalizado para RVT",
        "custom_prefix_for_dwg": "Prefijo personalizado para DWG",
        "results_title": "Seleccionar opciones",
        "results_button_name": "Guardar selección",
        "custom_prefix_rvt_value": "Elija un prefijo para los RVT",
        "custom_prefix_dwg_value": "Elija un prefijo para los DWG",
    },
    "de_de": {
        "set_type_ws": "Bearbeitungsbereich für Typ festlegen",
        "set_all": "Alle Verknüpfungen sammeln",
        "custom_prefix_for_rvt": "Benutzerdefiniertes Präfix für RVT",
        "custom_prefix_for_dwg": "Benutzerdefiniertes Präfix für DWG",
        "results_title": "Optionen auswählen",
        "results_button_name": "Auswahl speichern",
        "custom_prefix_rvt_value": "Wählen Sie ein Präfix für RVTs",
        "custom_prefix_dwg_value": "Wählen Sie ein Präfix für DWGs",
    },
}  # type: dict[str, dict[str, str]]
translations_dict = translations.get(pyrevit_locale, translations["en_us"])

opts = [
    MyOption(translations_dict["set_type_ws"], set_type_ws),
    MyOption(translations_dict["set_all"], set_all),
    MyOption(translations_dict["custom_prefix_for_rvt"], custom_prefix_for_rvt),
    MyOption(translations_dict["custom_prefix_for_dwg"], custom_prefix_for_dwg),
]

results = forms.SelectFromList.show(
    opts,
    multiselect=True,
    title=translations_dict["results_title"],
    button_name=translations_dict["results_button_name"],
    return_all=True,
    width=330,
    height=300,
)

if results:
    selected_items = {item.item: item.state for item in results}

    my_config.set_option("set_type_ws", selected_items.get(translations_dict["set_type_ws"], False))
    my_config.set_option("set_all", selected_items.get(translations_dict["set_all"], False))
    my_config.set_option("custom_prefix_for_rvt", selected_items.get(translations_dict["custom_prefix_for_rvt"], False))
    my_config.set_option("custom_prefix_for_dwg", selected_items.get(translations_dict["custom_prefix_for_dwg"], False))
    if selected_items.get(translations_dict["custom_prefix_for_rvt"], False):
        custom_prefix_dwg_value = my_config.get_option("custom_prefix_rvt_value", "ZL_RVT_")
        custom_prefix_dwg_value = forms.ask_for_string(
            default=custom_prefix_dwg_value,
            prompt=translations_dict["custom_prefix_rvt_value"]
        )
        my_config.set_option("custom_prefix_rvt_value", custom_prefix_dwg_value)
    if selected_items.get(translations_dict["custom_prefix_for_dwg"], False):
        custom_prefix_dwg_value = my_config.get_option("custom_prefix_dwg_value", "ZL_DWG_")
        custom_prefix_dwg_value = forms.ask_for_string(
            default=custom_prefix_dwg_value,
            prompt=translations_dict["custom_prefix_dwg_value"]
        )
        my_config.set_option("custom_prefix_dwg_value", custom_prefix_dwg_value)

    script.save_config()
    main()
