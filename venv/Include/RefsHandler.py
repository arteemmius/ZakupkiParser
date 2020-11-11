from selenium import webdriver
from parsel import Selector
from openpyxl import load_workbook
import re
import pandas as pd
import time

def checkNone(a):
    if a == None:
        a = ""
    return a


def append_df_to_excel(filename, df, sheet_name='Sheet1', startrow=None,
                       truncate_sheet=False,
                       **to_excel_kwargs):
    """
    Append a DataFrame [df] to existing Excel file [filename]
    into [sheet_name] Sheet.
    If [filename] doesn't exist, then this function will create it.

    Parameters:
      filename : File path or existing ExcelWriter
                 (Example: '/path/to/file.xlsx')
      df : dataframe to save to workbook
      sheet_name : Name of sheet which will contain DataFrame.
                   (default: 'Sheet1')
      startrow : upper left cell row to dump data frame.
                 Per default (startrow=None) calculate the last row
                 in the existing DF and write to the next row...
      truncate_sheet : truncate (remove and recreate) [sheet_name]
                       before writing DataFrame to Excel file
      to_excel_kwargs : arguments which will be passed to `DataFrame.to_excel()`
                        [can be dictionary]

    Returns: None
    """

    # ignore [engine] parameter if it was passed
    if 'engine' in to_excel_kwargs:
        to_excel_kwargs.pop('engine')

    writer = pd.ExcelWriter(filename, engine='openpyxl')

    try:
        # try to open an existing workbook
        writer.book = load_workbook(filename)

        # get the last row in the existing Excel sheet
        # if it was not specified explicitly
        if startrow is None and sheet_name in writer.book.sheetnames:
            startrow = writer.book[sheet_name].max_row

        # truncate sheet
        if truncate_sheet and sheet_name in writer.book.sheetnames:
            # index of [sheet_name] sheet
            idx = writer.book.sheetnames.index(sheet_name)
            # remove [sheet_name]
            writer.book.remove(writer.book.worksheets[idx])
            # create an empty sheet [sheet_name] using old index
            writer.book.create_sheet(sheet_name, idx)

        # copy existing sheets
        writer.sheets = {ws.title:ws for ws in writer.book.worksheets}
    except FileNotFoundError:
        # file does not exist yet, we will create it
        pass

    if startrow is None:
        startrow = 0

    # write out the new sheet
    df.to_excel(writer, sheet_name, startrow=startrow, **to_excel_kwargs)

    # save the workbook
    writer.save()


def CreateCsvData(finalList):
    df = pd.DataFrame(
        {'Реестровый номер контракта': finalList["number"], 'Реквизиты заказа': finalList["contact"],
         'Способ определения поставщика (подрядчика, исполнителя)': finalList["methodInit"], 'Дата подведения результатов определения поставщика (подрядчика, исполнителя)': finalList["dateRes"],
         'Дата размещения (по местному времени)': finalList["dateRealise"], 'Основание заключения контракта с единственным поставщиком': finalList["proofUnic"],
         'Полное наименование заказчика': finalList["fullName"], 'Уровень бюджета': finalList["cashLevel"], 'Дата заключения контракта': finalList["dateDoc"],
         'Цена контракта': finalList["costDoc"], 'Дата начала исполнения контракта': finalList["startdDt"], 'Дата окончания исполнения контракта': finalList["endDt"],
         'ОРГАНИЗАЦИЯ': finalList["orgDiller"], 'Извещение о проведении электронного аукциона': finalList["listAuction"],
         'Дата и время начала подачи заявок': finalList["startSet"], 'Дата и время окончания подачи заявок': finalList["endSet"],
         'Дата окончания срока рассмотрения первых частей заявок участников': finalList["endPeriod"], ' Дата проведения аукциона в электронной форме': finalList["dateAuction"],
         'Начальная (максимальная) цена контракта': finalList["maxCost"], 'УЧАСТНИК(И), С КОТОРЫМ ПЛАНИРУЕТСЯ ЗАКЛЮЧИТЬ КОНТРАКТ': finalList["peopleChoose"],
         })
    #writer = pd.ExcelWriter('parse_res.xlsx', engine='openpyxl', encoding='ISO-8859-1', mode='a')
    #df.to_excel('parse_res.xlsx', encoding='ISO-8859-1', index=False)
    append_df_to_excel('parse_res.xlsx', df, sheet_name='Sheet1', index=False, header=None)


driver = webdriver.Chrome("c:/utils/chromedriver_win32/chromedriver.exe")
# размер окна
driver.set_window_size(1500, driver.get_window_size()['height'])
orgAtrs = dict()
orgsData = list()

f = open('orgsData.txt')
for line in f:
    #init keys
    orgAtrs["number"] = []
    orgAtrs["contact"] = []
    orgAtrs["methodInit"] = []
    orgAtrs["dateRes"] = []
    orgAtrs["dateRealise"] = []
    orgAtrs["proofUnic"] = []
    orgAtrs["fullName"] = []
    orgAtrs["cashLevel"] = []
    orgAtrs["dateDoc"] = []
    orgAtrs["costDoc"] = []
    orgAtrs["startdDt"] = []
    orgAtrs["endDt"] = []
    orgAtrs["orgDiller"] = []
    orgAtrs["listAuction"] = []
    orgAtrs["startSet"] = []
    orgAtrs["endSet"] = []
    orgAtrs["endPeriod"] = []
    orgAtrs["dateAuction"] = []
    orgAtrs["maxCost"] = []
    orgAtrs["peopleChoose"] = []

    orgsData = line.split(";")
    orgAtrs["number"].append(orgsData[0])
    orgAtrs["contact"].append(orgsData[1])
    time.sleep(3)
    driver.get(orgsData[2])
    sel = Selector(text=driver.find_element_by_xpath("//*").get_attribute("outerHTML"))
    methodInit = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][1]//tr/td[1][./text()='Способ определения поставщика (подрядчика, исполнителя)']/../td[2]/text()").extract_first()
    dateRes = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][1]//tr/td[1][./text()='Дата подведения результатов определения поставщика (подрядчика, исполнителя)']/../td[2]/text()").extract_first()
    dateRealise = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][1]//tr/td[1][./text()='Дата размещения (по местному времени)']/../td[2]/text()").extract_first()
    proofUnic = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][1]//tr/td[1][./text()='Основание заключения контракта с единственным поставщиком']/../td[2]/text()").extract_first()

    orgAtrs["methodInit"].append(checkNone(methodInit.replace("  ", "").replace("\n", "")))
    orgAtrs["dateRes"].append(checkNone(dateRes))
    orgAtrs["dateRealise"].append(checkNone(dateRealise))
    orgAtrs["proofUnic"].append(checkNone(proofUnic))

    if sel.xpath("//div[@class='contentTabBoxBlock']//h2[@class='noticeBoxH2'][2]/text()").extract_first() != "Информация об изменении контракта":
        j = 0
    else:
        j = 1
    fullName = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][" + str(j + 2) + "]//tr/td[1][./text()='Полное наименование заказчика']/../td[2]/a/text()").extract_first()
    cashLevel = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][" + str(j + 2) + "]//tr/td[1][./text()='Уровень бюджета']/../td[2]/text()").extract_first()
    dateDoc = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][" + str(j + 3) + "]//tr/td[1][./text()='Дата заключения контракта']/../td[2]/text()").extract_first()
    costDoc = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][" + str(j + 3) + "]//tr/td[1][./text()='Цена контракта']/../td[2]/text()").extract_first()
    startdDt = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][" + str(j + 3) + "]//tr/td[1][./text()='Дата начала исполнения контракта']/../td[2]/text()").extract_first()
    endDt = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][" + str(j + 3) + "]//tr/td[1][./text()='Дата окончания исполнения контракта']/../td[2]/text()").extract_first()
    orgDiller = sel.xpath(
        "//div[@class='noticeTabBoxWrapper'][" + str(j + 4) + "]//tr[3]/td/text()[1]").extract_first() + " "
    if sel.xpath("//div[@class='noticeTabBoxWrapper'][" + str(j + 4) + "]//tr[3]/td/text()[2]").extract_first() != None:
        orgDiller = orgDiller + sel.xpath("//div[@class='noticeTabBoxWrapper'][" + str(j + 4) + "]//tr[3]/td/text()[2]").extract_first()

    orgAtrs["fullName"].append(checkNone(fullName).replace("\xa0", "").replace("  ", "").replace("\n", ""))
    orgAtrs["cashLevel"].append(checkNone(cashLevel).replace("\xa0", ""))
    orgAtrs["dateDoc"].append(checkNone(dateDoc).replace("\xa0", ""))
    orgAtrs["costDoc"].append(checkNone(costDoc).replace("\xa0", ""))
    orgAtrs["startdDt"].append(checkNone(startdDt).replace("\xa0", ""))
    orgAtrs["endDt"].append(checkNone(endDt).replace("\xa0", ""))
    orgAtrs["orgDiller"].append(checkNone(orgDiller).replace("\xa0", "").replace("  ", "").replace("\n", ""))
    if orgsData[1] != "":
        time.sleep(3)
        driver.get(orgsData[3].replace("\n", ""))
        sel = Selector(text=driver.find_element_by_xpath("//*").get_attribute("outerHTML"))
        listAuction = sel.xpath("//div[@class='padBtm20 contentHeadingWrapper']/a/@href").extract_first()
        startSet = sel.xpath(
            "//div[@class='noticeTabBoxWrapper'][3]//tr/td[1][./text()='Дата и время начала подачи заявок']/../td[2]/text()").extract_first()
        endSet = sel.xpath(
            "//div[@class='noticeTabBoxWrapper'][3]//tr/td[1][./text()='Дата и время окончания подачи заявок']/../td[2]/text()").extract_first()
        endPeriod = sel.xpath(
            "//div[@class='noticeTabBoxWrapper'][3]//tr/td[1][./text()='Дата окончания срока рассмотрения первых частей заявок участников']/../td[2]/text()").extract_first()
        dateAuction = sel.xpath(
            "//div[@class='noticeTabBoxWrapper'][3]//tr/td[1][./text()='Дата проведения аукциона в электронной форме']/../td[2]/text()").extract_first()
        maxCost = sel.xpath(
            "//div[@class='noticeTabBoxWrapper'][4]//tr/td[1][./text()='Начальная (максимальная) цена контракта']/../td[2]/text()").extract_first()
        if checkNone(listAuction) != "":
            orgAtrs["listAuction"].append("http://zakupki.gov.ru" + checkNone(listAuction))
        else:
            orgAtrs["listAuction"].append("")
        orgAtrs["startSet"].append(checkNone(startSet))
        orgAtrs["endSet"].append(checkNone(endSet))
        orgAtrs["endPeriod"].append(checkNone(endPeriod))
        orgAtrs["dateAuction"].append(checkNone(dateAuction).replace("\n", " ").replace("  ", ""))
        orgAtrs["maxCost"].append(checkNone(maxCost).replace("\xa0", ""))
        time.sleep(2)
        driver.find_element_by_xpath("//td[@tab='SUPPLIER_RESULTS']").click()
        sel = Selector(text=driver.find_element_by_xpath("//*").get_attribute("outerHTML"))
        peopleChooseFull = ""
        i = 0
        while 1:
            if i == 0:
                peopleChoose = sel.xpath(
                    "//table[@class='borderSpacingSeparate2px']//tr[" + str(i + 3) + "]//td[3]/text()").extract_first()
            else:
                peopleChoose = sel.xpath(
                    "//table[@class='borderSpacingSeparate2px']//tr[" + str(i + 3) + "]//td[1]/text()").extract_first()

            if peopleChoose != None:
                peopleChooseFull = peopleChooseFull + peopleChoose + ";"
            else:
                break
            i = i + 1
        orgAtrs["peopleChoose"].append(checkNone(peopleChooseFull).replace("\n", " ").replace("  ", "")
                                       .replace(" По окончании срока подачи заявок подана только одна заявка. Такая заявка признана соответствующей требованиям Федерального закона № 44-ФЗ и документации об аукционе. Электронный аукцион признан несостоявшимся по основанию, предусмотренному ч. 1 ст. 71 Федерального закона № 44-ФЗ ; ", "").
                                       replace(" В течение десяти минут после начала проведения электронного аукциона было подано единственное предложение о цене контракта. По результатам рассмотрения второй части такой заявки принято решение о ее соответствии требованиям, установленным документацией об электронном аукционе (ч.13 ст.69 Закона № 44-ФЗ). ;","").
                                       replace(" Принято решение о признании только одного участника закупки, подавшего заявку на участие в электронном аукционе, его участником. По результатам рассмотрения заявки единственного участника, данная заявка признана соответствующей требованиям Федерального закона № 44-ФЗ и документации об аукционе. (п. 3 ч.2 ст. 71) ;", ""))
    else:
        orgAtrs["listAuction"].append("")
        orgAtrs["startSet"].append("")
        orgAtrs["endSet"].append("")
        orgAtrs["endPeriod"].append("")
        orgAtrs["dateAuction"].append("")
        orgAtrs["maxCost"].append("")
        orgAtrs["peopleChoose"].append("")

    CreateCsvData(orgAtrs)
    orgAtrs.clear()
    print("processed: " + line)
driver.close()
