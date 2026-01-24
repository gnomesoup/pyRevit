# -*- coding: UTF-8 -*-
import io
import json
import os
from pyrevit import script, forms
from pyrevit.userconfig import user_config
from script import main


def get_translations(script_folder, script_type):
    # type: (str, str) -> dict[str, dict[str, str | list]]
    """Loads translations for a specific script type from a JSON file."""
    json_path = os.path.join(script_folder, 'translations.json')
    with io.open(json_path, 'r', encoding='utf-8') as file:
        all_translations = json.load(file)
    return all_translations[script_type]


class MyOption(forms.TemplateListItem):
    @property
    def name(self):
        return str(self.item)


my_config = script.get_config()

set_type_ws = my_config.get_option("set_type_ws", False)
set_all = my_config.get_option("set_all", False)
custom_prefix_for_rvt = my_config.get_option("custom_prefix_for_rvt", False)
custom_prefix_for_dwg = my_config.get_option("custom_prefix_for_dwg", False)

pyrevit_locale = user_config.user_locale  # type: str
translations = get_translations(os.path.dirname(__file__), "config")
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
        custom_prefix_value = my_config.get_option("custom_prefix_rvt_value", "ZL_RVT_")
        custom_prefix_value = forms.ask_for_string(
            default=custom_prefix_value,
            prompt=translations_dict["custom_prefix_rvt_value"]
        )
        my_config.set_option("custom_prefix_rvt_value", custom_prefix_value)
    if selected_items.get(translations_dict["custom_prefix_for_dwg"], False):
        custom_prefix_value = my_config.get_option("custom_prefix_dwg_value", "ZL_DWG_")
        custom_prefix_value = forms.ask_for_string(
            default=custom_prefix_value,
            prompt=translations_dict["custom_prefix_dwg_value"]
        )
        my_config.set_option("custom_prefix_dwg_value", custom_prefix_value)

    script.save_config()
    main()
