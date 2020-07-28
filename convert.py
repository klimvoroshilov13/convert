# Created by N.Kazakov ver 1.00

# python imports
import sys
import re
import string
import collections
from xml.dom import minidom
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL


class Helpers:

    @staticmethod
    def showError(parentwin, error):
        box = parentwin.getToolkit().createMessageBox(
            parentwin, ERRORBOX, BUTTONS_OK, "Ошибка", error)
        box.execute()

    @staticmethod
    def getMonthNum(month):
        months = {
            "янв": "01",
            "фев": "02",
            "мар": "03",
            "апр": "04",
            "май": "05",
            "июн": "06",
            "июл": "07",
            "авг": "08",
            "сен": "09",
            "окт": "10",
            "ноя": "11",
            "дек": "12"}
        month = month[0:3]
        return months.get(month, False)

    @staticmethod
    def getData(parentwin, headDocument, string, cell, step=2):
        if re.search(headDocument, string):
            string = string[len(headDocument) + step:len(string)]
        else:
            error = "Error getting information in cell " + cell
            Helpers.showError(parentwin, error)
            try:
                sys.exit()
            except SystemExit:
                return None
        return string

    @staticmethod
    def createTag(name: str = None, text: str = None, attributes: dict = None, *, cdata: bool = False):
        doc = minidom.Document()
        if name is None:
            return doc
        tag = doc.createElement(name)
        if text is not None:
            if cdata is True:
                tag.appendChild(doc.createCDATASection(text))
            else:
                tag.appendChild(doc.createTextNode(text))
        if attributes is not None:
            for k, v in attributes.items():
                tag.setAttribute(k, str(v))
        return tag


def main(*args):
    root = {"Файл": {"ИдФайл": "ON_NSCHFDOPPR_", "ВерсФорм": "1.00", "ВерсПрог": "convert"}}
    branchs = {
        "СвУчДокОбор": {"ИдПол": "", "ИдОтпр": ""},
        "Документ": {"Функция": "СЧФ", "НаимЭконСубСост": "", "ДатаИнфПр": "", "ВремИнфПр": "00.00.00", "КНД": "1115131"},
        "СвСчФакт": {"КодОКВ": "", "ДатаСчФ": "", "НомерСчФ": ""},
        "СвПрод": None, "ИдСв": None, "Адрес": None,
        "СвПокуп": None,
        "ТаблСчФакт": None,
        "СведТов": {
            "НомСтр": "", "НаимТов": "", "ОКЕИ_Тов": "", "НаимЕдИзм": "", "КолТов": "",
            "ЦенаТов": "", "СтТовБезНДС": "", "БезАкциз": "", "НалСт": "", "СумНал": "", "СтТовУчНал": ""},
        "Акциз": None,
        "СумНал": None,
        "ВсегоОпл": {"СтТовУчНалВсего": "", "СтТовБезНДСВсего": ""},
        "СумНалВсего": None,
        "Подписант": None,
        "ЮЛ": {"ИННЮЛ": "", "Должн": ""}}
    leafs = {
        "СвОЭДОтпр": {"НаимОрг": "Наименование ЮЛ Оператора", "ИдЭДО": "", "ИННЮЛ": ""},
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
        "НомСтр", "НаимТов", "ОКЕИ_Тов", "НаимЕдИзм", "КолТов",
        "ЦенаТов", "СтТовБезНДС", "БезАкциз", "НалСт", "СумНал", "СтТовУчНал"]
    cellsDocument = ["B", "F", "I", "K", "M", "O", "R", "S", "V", "X"]

        # get the doc from the scripting context.which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    sheets = model.Sheets
    fullPath = model.URL[8:len(model.URL)]
    path = ""
    count = 0
    for i in fullPath:
        if i != "%" and count == 0:
            path += i
        else:
            count += 1
    parentwin = model.CurrentController.Frame.ContainerWindow
    sheet = sheets[0]
    # Cell B2
    string = Helpers.getData(
        parentwin, headDocument[0],
        sheet.getCellRangeByName("B2").String,
        "B2",
        3)
    space = 0
    day = ""
    month = ""
    year = ""
    for i in string:
        if i != " ":
            if space == 0:
                branchs["СвСчФакт"]["НомерСчФ"] += i
            elif space == 2:
                day += i
            elif space == 3:
                month += i
            elif space == 4:
                year += i
            elif space > 5:
                error = "Error incorrect cell B2 entry "
                Helpers.showError(parentwin, error)
                try:
                    sys.exit()
                except SystemExit:
                    return None
        else:
            space += 1
    monthNum = Helpers.getMonthNum(month)
    if monthNum:
        date = day + "." + monthNum + "." + year
        branchs["Документ"]["ДатаИнфПр"] = date
        branchs["СвСчФакт"]["ДатаСчФ"] = date
    else:
        error = "Error incorrect date in cell B2"
        Helpers.showError(parentwin, error)
        try:
            sys.exit()
        except SystemExit:
            return None
    # Cell B4
    string = Helpers.getData(
        parentwin, headDocument[1],
        sheet.getCellRangeByName("B4").String, "B4")
    leafs["СвЮЛПрод"]["НаимОрг"] = string
    branchs["Документ"]["НаимЭконСубСост"] = string
    # Cell B5
    string = Helpers.getData(
        parentwin, headDocument[2],
        sheet.getCellRangeByName("B5").String, "B5")
    leafs["АдрИнфПрод"]["АдрТекст"] = string
    # Cell B6
    string = Helpers.getData(
        parentwin,
        headDocument[3],
        sheet.getCellRangeByName("B6").String, "B6")
    slash = 0
    for i in string:
        if i != "/":
            if slash == 0:
                leafs["СвЮЛПрод"]["ИННЮЛ"] += i
            elif slash == 1:
                leafs["СвЮЛПрод"]["КПП"] += i
        else:
            slash += 1
    # Cell B10
    string = Helpers.getData(
        parentwin, headDocument[4],
        sheet.getCellRangeByName("B10").String,
        "B10")
    leafs["СвЮЛПокуп"]["НаимОрг"] = string
    # Cell B11
    string = Helpers.getData(
        parentwin, headDocument[5],
        sheet.getCellRangeByName("B11").String,
        "B11")
    leafs["АдрИнфПокуп"]["АдрТекст"] = string
    # Cell B12
    string = Helpers.getData(
        parentwin,
        headDocument[6],
        sheet.getCellRangeByName("B12").String,
        "B12")
    slash = 0
    for i in string:
        if i != "/":
            if slash == 0:
                leafs["СвЮЛПокуп"]["ИННЮЛ"] += i
            elif slash == 1:
                leafs["СвЮЛПокуп"]["КПП"] += i
        else:
            slash += 1
    # Cell B13
    string = Helpers.getData(
        parentwin, headDocument[7],
        sheet.getCellRangeByName("B13").String,
        "B13",
        38)
    leafs["АдрИнфПрод"]["КодСтр"] = string
    leafs["АдрИнфПокуп"]["КодСтр"] = string
    branchs["СвСчФакт"]["КодОКВ"] = string
    # Row 18
    j = 1
    for i in [*cellsDocument]:
        if sheet.getCellRangeByName(i + "18").String:
            string = sheet.getCellRangeByName(i + "18").String
            string = string.replace("\xa0", "")
            branchs["СведТов"][bodyDocument[j]] = string.replace(",", ".")
            j += 1
        else:
            error = "Error incorrect date in cell " + i + "18"
            Helpers.showError(parentwin, error)
            try:
                sys.exit()
            except SystemExit:
                return None
    data["БезАкциз"] = branchs["СведТов"]["БезАкциз"]
    data["СумНал"] = float(branchs["СведТов"]["СумНал"])

    doc = Helpers.createTag()

    fileDoc = root["Файл"]
    fileDoc = Helpers.createTag(root.popitem()[0], attributes=fileDoc)
    doc.appendChild(fileDoc)

    keysBranchs = list(branchs.keys())
    keysLeafs = list(leafs.keys())

    infoDocTurn = branchs["СвУчДокОбор"]
    infoDocTurn = Helpers.createTag(keysBranchs[0], attributes=infoDocTurn)
    fileDoc.appendChild(infoDocTurn)

    infoDigSent = leafs["СвОЭДОтпр"]
    infoDigSent = Helpers.createTag(keysLeafs[0], attributes=infoDigSent)
    infoDocTurn.appendChild(infoDigSent)

    infoDoc = branchs["Документ"]
    infoDoc = Helpers.createTag(keysBranchs[1], attributes=infoDoc)
    fileDoc.appendChild(infoDoc)

    infoInvoice = branchs["СвСчФакт"]
    infoInvoice = Helpers.createTag(keysBranchs[2], attributes=infoInvoice)
    infoDoc.appendChild(infoInvoice)

    infoSeller = branchs["СвПрод"]
    infoSeller = Helpers.createTag(keysBranchs[3], attributes=infoSeller)
    infoInvoice.appendChild(infoSeller)

    infoId = branchs["ИдСв"]
    infoId = Helpers.createTag(keysBranchs[4], attributes=infoId)
    infoSeller.appendChild(infoId)

    xmlData = doc.toprettyxml(indent=" ", newl="\n", encoding="windows-1251")
    dateDoc = branchs["СвСчФакт"]["ДатаСчФ"]
    numDoc = " №" + branchs["СвСчФакт"]["НомерСчФ"] + " от " + dateDoc
    nameDoc = branchs["Документ"]["Функция"] + numDoc
    try:
        with open(path + nameDoc + ".xml", "w") as f:
            f.write(xmlData.decode("windows-1251"))
    except IOError:
        error = "Error opening file"
        Helpers.showError(parentwin, error)

    return None
