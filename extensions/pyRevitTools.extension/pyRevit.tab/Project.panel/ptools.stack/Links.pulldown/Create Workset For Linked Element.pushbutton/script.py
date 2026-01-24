# -*- coding: UTF-8 -*-
import io
import json
import os
from pyrevit import revit, DB
from pyrevit import script
from pyrevit.forms import alert
from pyrevit.userconfig import user_config
from pyrevit import HOST_APP


def get_translations(script_folder, script_type):
    # type: (str, str) -> dict[str, dict[str, str | list]]
    """Loads translations for a specific script type from a JSON file."""
    json_path = os.path.join(script_folder, 'translations.json')
    with io.open(json_path, 'r', encoding='utf-8') as file:
        all_translations = json.load(file)
    return all_translations[script_type]


doc = HOST_APP.doc
logger = script.get_logger()
pyrevit_locale = user_config.user_locale  # type: str
translations = get_translations(script.get_script_path(), "script")
translations_dict = translations.get(pyrevit_locale, translations["en_us"])


def main():
    """Main function to create worksets for linked elements."""
    my_config = script.get_config()

    set_type_ws = my_config.get_option("set_type_ws", False)
    set_all = my_config.get_option("set_all", False)
    custom_prefix_for_rvt = my_config.get_option("custom_prefix_for_rvt", False)
    custom_prefix_for_dwg = my_config.get_option("custom_prefix_for_dwg", False)

    if not set_all:
        selection = revit.get_selection()
    else:
        selection = (
            DB.FilteredElementCollector(doc)
            .WhereElementIsNotElementType()
            .WherePasses(DB.LogicalOrFilter([
                DB.ElementClassFilter(DB.RevitLinkInstance),
                DB.ElementClassFilter(DB.ImportInstance)
            ]))
            .ToElements()
        )

    if len(selection) > 0:
        enable_worksharing = alert(
            translations_dict["enable_worksharing"],
            options=translations_dict["enable_worksharing_options"],
            warn_icon=False
        )  # type: str
        if not enable_worksharing:
            script.exit()
        if (
            enable_worksharing == translations_dict["enable_worksharing_options"][0]
            and not doc.IsWorkshared
            and doc.CanEnableWorksharing
        ):
            doc.EnableWorksharing("Shared Levels and Grids", "Workset1")
        else:
            alert(
                translations_dict["enable_worksharing_no"],
                title=translations_dict["enable_worksharing_no_title"],
                exitscript=True
            )
        
        with revit.Transaction(translations_dict["transaction_name"]):
            for el in selection:
                linked_model_name = ""
                if isinstance(el, DB.RevitLinkInstance):
                    prefix_for_rvt_value = "RVT_"
                    if custom_prefix_for_rvt:
                        prefix_for_rvt_value = my_config.get_option(
                            "custom_prefix_rvt_value", prefix_for_rvt_value
                        )
                    linked_model_name = (
                        prefix_for_rvt_value + el.Name.split(":")[0].split(".rvt")[0]
                    )
                elif isinstance(el, DB.ImportInstance):
                    prefix_for_dwg_value = "DWG_"
                    if custom_prefix_for_dwg:
                        prefix_for_dwg_value = my_config.get_option(
                            "custom_prefix_dwg_value", prefix_for_dwg_value
                        )
                    linked_model_name = (
                        prefix_for_dwg_value
                        + el.get_Parameter(DB.BuiltInParameter.IMPORT_SYMBOL_NAME)
                        .AsString()
                        .split(".dwg")[0]
                    )
                if linked_model_name:
                    try:
                        user_worksets = (
                            DB.FilteredWorksetCollector(doc)
                            .OfKind(DB.WorksetKind.UserWorkset)
                            .ToWorksets()
                        )
                        existing_ws = None
                        for ws in user_worksets:
                            if ws.Name == linked_model_name:
                                existing_ws = ws
                                break
                        if existing_ws:
                            workset = existing_ws
                        else:
                            workset = DB.Workset.Create(doc, linked_model_name)
                        worksetParam = el.get_Parameter(
                            DB.BuiltInParameter.ELEM_PARTITION_PARAM
                        )
                        success = False
                        if not worksetParam.IsReadOnly:
                            worksetParam.Set(workset.Id.IntegerValue)
                            success = True
                        else:
                            logger.error("Instance Workset Parameter is read-only")
                        if set_type_ws:
                            type_id = el.GetTypeId()
                            type_el = doc.GetElement(type_id)
                            type_workset_param = type_el.get_Parameter(
                                DB.BuiltInParameter.ELEM_PARTITION_PARAM
                            )
                            if not type_workset_param.IsReadOnly:
                                type_workset_param.Set(workset.Id.IntegerValue)
                                success = True
                            else:
                                logger.error("Type Workset Parameter is read-only")
                        if not success and not existing_ws:
                            workset_table = doc.GetWorksetTable()
                            workset_table.DeleteWorkset(
                                doc, workset.Id, DB.DeleteWorksetSettings()
                            )
                    except Exception as e:
                        logger.error(
                            "Error setting Workset for: {}\nError: {}".format(
                                linked_model_name, e
                            )
                        )
    else:
        if set_all:
            alert(translations_dict["set_all_no_links"])
        else:
            alert(translations_dict["set_all_select_at_least_one"])


if __name__ == "__main__":
    main()
