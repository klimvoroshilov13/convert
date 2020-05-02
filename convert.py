# Created by N.Kazakov ver 1.00

# python imports
import sys
import re
from xml.dom import minidom

def main(*args):
    # get the doc from the scripting context.which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    sheets = model.Sheets
    sheet = sheets[0]
    root = {Файл : None, ВерсФорм : "1.00", ВерсПрог : "convert", ИдФайл : "ON_NSCHFDOPPR_"}
    branchs = {
        СвУчДокОбор : None, ИдПол : "", ИдОтпр : "",
        Документ : None, Функция : "СЧФ", НаимЭконСубСост : "", ДатаИнфПр : "", ВремИнфПр : "00.00.00", КНД : "1115131",
        СвСчФакт : None, КодОКВ : "", ДатаСчФ : "", НомерСчФ : "",
        СвПрод : None,
        СвПокуп : None,
        ИдСв : None,
        Адрес : None,
        ТаблСчФакт : None,
        СведТов : None, СтТовУчНал : "", НалСт : "", СтТовБезНДС : "", ЦенаТов : "", КолТов : "", ОКЕИ_Тов : "", НаимТов : "", НомСтр : "",
        Акциз : None,
        СумНал : None,
        ВсегоОпл : None, СтТовУчНалВсего : "", СтТовБезНДСВсего : "",
        СумНалВсего : None,
        Подписант : None,
        ЮЛ : None, ИННЮЛ : "", Должн : ""}
    leafs = {СвОЭДОтпр : None, ИдЭДО : "", ИННЮЛ : "", НаимОрг : "Наименование ЮЛ Оператора",
        СвЮЛУч : None, ИННЮЛ : "", НаимОрг : "", КПП : "",
        АдрИнф : None, АдрТекст : "", КодСтр : "",
        ДопСведТов : None, НаимЕдИзм : "",
        ФИО : None, Отчество : "", Имя : "", Фамилия : ""}
    data = {БезАкциз : "", СумНал : 0}

    return None