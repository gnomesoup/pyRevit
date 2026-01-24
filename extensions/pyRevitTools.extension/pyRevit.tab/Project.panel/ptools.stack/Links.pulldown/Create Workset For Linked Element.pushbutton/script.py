# -*- coding: UTF-8 -*-
from pyrevit import revit, DB
from pyrevit import script
from pyrevit.forms import alert
from pyrevit.userconfig import user_config
from pyrevit import HOST_APP

doc = HOST_APP.doc
logger = script.get_logger()

pyrevit_locale = user_config.user_locale  # type: str
# Translation dictionaries
translations = {
    "en_us": {
        "enable_worksharing": "The document doesn't have worksharing enabled.\nEnable it?",
        "enable_worksharing_options": (
            "Yes",
            "No"
        ),
        "enable_worksharing_no": "The script cannot run in a document without worksharing.",
        "enable_worksharing_no_title": "The script has stopped",
        "transaction_name": "Create Workset(s) for linked model(s)",
        "set_all_no_links": "No links found in the document.",
        "set_all_select_at_least_one": "At least one linked element must be selected.",
    },
    "fr_fr": {
        "enable_worksharing": "Le partage de projet n'est pas activé pour ce document.\nL'activer ?",
        "enable_worksharing_options": (
            "Oui",
            "Non"
        ),
        "enable_worksharing_no": "Le script ne peut pas s'exécuter dans un document sans partage de projet.",
        "enable_worksharing_no_title": "Le script s'est arrêté",
        "transaction_name": "Créer des sous-projets pour les modèles liés",
        "set_all_no_links": "Aucun lien trouvé dans le document.",
        "set_all_select_at_least_one": "Au moins un élément lié doit être sélectionné.",
    },
    "ru": {
        "enable_worksharing": "Документ без совместной работы.\nВключить её?",
        "enable_worksharing_options": (
            "Да",
            "Нет"
        ),
        "enable_worksharing_no": "Скрипт не может работать в документе без совместной работы.",
        "enable_worksharing_no_title": "Скрипт остановлен",
        "transaction_name": "Создание рабочих наборов для связанных моделей",
        "set_all_no_links": "В документе нет связанных файлов.",
        "set_all_select_at_least_one": "Необходимо выбрать хотя бы один элемент связи.",
    },
    "chinese_s": {
        "enable_worksharing": "文档未启用工作共享。\n要启用它吗？",
        "enable_worksharing_options": (
            "是",
            "否"
        ),
        "enable_worksharing_no": "该脚本无法在没有工作共享的文档中运行。",
        "enable_worksharing_no_title": "脚本已停止",
        "transaction_name": "为链接模型创建工作集",
        "set_all_no_links": "在文档中找不到链接。",
        "set_all_select_at_least_one": "必须至少选择一个链接元素。",
    },
    "es_es": {
        "enable_worksharing": "El documento no tiene activada la compartición de proyecto.\n¿Activarla?",
        "enable_worksharing_options": (
            "Sí",
            "No"
        ),
        "enable_worksharing_no": "El script no puede ejecutarse en un documento sin compartición de proyecto.",
        "enable_worksharing_no_title": "El script se ha detenido",
        "transaction_name": "Crear subproyectos para modelos vinculados",
        "set_all_no_links": "No se encontraron vínculos en el documento.",
        "set_all_select_at_least_one": "Se debe seleccionar al menos un elemento vinculado.",
    },
    "de_de": {
        "enable_worksharing": "Die Bearbeitungsbereiche sind für das Dokument nicht aktiviert.\nAktivieren?",
        "enable_worksharing_options": (
            "Ja",
            "Nein"
        ),
        "enable_worksharing_no": "Das Skript kann nicht in einem Dokument ohne Bearbeitungsbereiche ausgeführt werden.",
        "enable_worksharing_no_title": "Das Skript wurde angehalten",
        "transaction_name": "Bearbeitungsbereiche für verknüpfte Modelle erstellen",
        "set_all_no_links": "Keine Verknüpfungen im Dokument gefunden.",
        "set_all_select_at_least_one": "Es muss mindestens ein verknüpftes Element ausgewählt werden.",
    },
}  # type: dict[str, dict[str, str | tuple]]
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
