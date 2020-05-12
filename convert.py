# Created by N.Kazakov ver 1.00

# python imports
import sys
import re
from xml.dom import minidom
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL


class Error:
    @staticmethod
    def showMessage(parentwin, error):
        box = parentwin.getToolkit().createMessageBox(
            parentwin, ERRORBOX, BUTTONS_OK, "Ошибка", error)
        box.execute()


def main(*args):
    root = {"Файл": {"ВерсФорм": "1.00", "ВерсПрог": "convert", "ИдФайл": "ON_NSCHFDOPPR_"}}
    branchs = {
        "СвУчДокОбор": {"ИдПол": "", "ИдОтпр": ""},
    "Документ": {"Функция": "СЧФ", "НаимЭконСубСост": "", "ДатаИнфПр": "", "ВремИнфПр": "00.00.00", "КНД": "1115131"},
        "СвСчФакт": {"КодОКВ": "", "ДатаСчФ": "", "НомерСчФ": ""},
        "СведТов": {
            "СтТовУчНал": "", "НалСт": "", "СтТовБезНДС": "", "ЦенаТов": "", "КолТов": "", "ОКЕИ_Тов": "", "НаимТов": "", "НомСтр": ""},
        "ВсегоОпл": {"СтТовУчНалВсего": "", "СтТовБезНДСВсего": ""},
        "ЮЛ": {"ИННЮЛ": "", "Должн": ""}}
    node = {
        "СвПрод": None,
        "СвПокуп": None,
        "ИдСв": None,
        "Адрес": None,
        "ТаблСчФакт": None,
        "Акциз": None,
        "СумНал": None,
        "СумНалВсего": None,
        "Подписант": None}
    leafs = {"СвОЭДОтпр": {"ИдЭДО": "", "ИННЮЛ": "", "НаимОрг": "Наименование ЮЛ Оператора"},
        "СвЮЛПрод": {"Name": "СвЮЛУч", "ИННЮЛ": "", "НаимОрг": "", "КПП": ""},
        "АдрИнфПрод": {"Name": "АдрИнф", "АдрТекст": "", "КодСтр": ""},
        "СвЮЛПокуп": {"Name": "СвЮЛУч", "ИННЮЛ": "", "НаимОрг": "", "КПП": ""},
        "АдрИнфПокуп": {"Name": "АдрИнф", "АдрТекст": "", "КодСтр": ""},
        "ДопСведТов": {"НаимЕдИзм": ""},
        "ФИО": {"Отчество": "", "Имя": "", "Фамилия": ""}}
    data = {"БезАкциз": "", "СумНал": 0}

    headDocument = [
        "Счет-фактура", "Продавец", "Адрес", "ИНН/КПП продавца", "Покупатель", "Адрес", "ИНН/КПП покупателя", "Валюта"]
    bodyDocument = [
        "Наименование товара", "код", "условное обозначение", "Коли-чество", "Цена", "без налога - всего", "сумма акциза",
        "Налоговая ставка", "Сумма налога", "с налогом - всего"]
    cellsDocument = ["B", "F", "I", "K", "M", "O", "R", "S", "V", "X"]

        # get the doc from the scripting context.which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    sheets = model.Sheets
    parentwin = model.CurrentController.Frame.ContainerWindow
    sheet = sheets[0]
    if re.search(headDocument[0], sheet.getCellRangeByName("B2").String):
        string = sheet.getCellRangeByName("B2").String
        string = string[string.find("№") + 2:len(string)]
        space = 0
        date = ""
        for i in string:
            if i != " ":
                if space == 0:
                    branchs["СвСчФакт"]["НомерСчФ"] += i
                elif space >= 2:
                    date += i
            else:
                space += 1
    else:
        error = "Error getting information in cell B2"
        Error.showMessage(parentwin, error)
        try:
            sys.exit()
        except SystemExit:
            return None

    for i in [*bodyDocument]:
        pass

    return None
