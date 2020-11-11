from selenium import webdriver
from parsel import Selector
import re
import time

driver = webdriver.Chrome("c:/utils/chromedriver_win32/chromedriver.exe")
# размер окна
driver.set_window_size(1500, driver.get_window_size()['height'])
# get html by url вызываем метод get и передаем ему url сайта

driver.get("http://zakupki.gov.ru/epz/contract/extendedsearch/results.html?morphology=on&openMode=USE_DEFAULT_PARAMS&pageNumber=1&sortDirection=false&recordsPerPage=_10&sortBy=PO_DATE_OBNOVLENIJA&fz44=on&contractPriceFrom=50000&contractPriceTo=200000000000&advancePercentFrom=hint&advancePercentTo=hint&contractStageList_1=on&contractStageList=1&regionDeleted=false&contractDateFrom=01.01.2014&budgetaryFunds=on&placingWayForContractList_4=on&placingWayForContractList_10=on&placingWayForContractList=4%2C10&okpd2IdsWithNested=on&okpd2Ids=8875284&okpd2IdsCodes=21.20.1&classifiersMpGroupId=0")
time.sleep(2)
for j in range(0, 100):
    sel = Selector(text=driver.find_element_by_xpath("//*").get_attribute("outerHTML"))
    numberList = list()
    numberList = sel.xpath("//td[@class='descriptTenderTd']/dl/dt/a/text()").extract()
    contactsList = list()
    refNumberList = list()
    refNumberList = sel.xpath("//td[@class='descriptTenderTd']/dl/dt/a/@href").extract()
    refContactsList = list()
    for k in range(1, 11):
        ref_i = sel.xpath("//div[@class='registerBox registerBoxBank margBtm20'][" + str(k) + "]//dd[@class='additionalDescriptionList']//a[@class='displayInlineBlockUsual widthAutoUsual']/@href").extract_first()
        if ref_i == None:
            ref_i = ""
        else:
            ref_i = "http://zakupki.gov.ru" + ref_i
        refContactsList.append(ref_i)
        name_i = sel.xpath("//div[@class='registerBox registerBoxBank margBtm20'][" + str(k) + "]//dd[@class='additionalDescriptionList']//a[@class='displayInlineBlockUsual widthAutoUsual']/text()").extract_first()
        if name_i == None:
            name_i = ""
        contactsList.append(name_i)

    f = open('orgsData.txt', 'a')
    for i in range(0, len(numberList)):
        f.write(numberList[i] + ";")
        f.write(contactsList[i].replace("  ", "").replace("\n", "") + ";")
        f.write("http://zakupki.gov.ru" + refNumberList[i] + ";")
        f.write(refContactsList[i])
        f.write("\n")

    numberList.clear()
    contactsList.clear()
    refContactsList.clear()
    refNumberList.clear()
    try:
        driver.find_element_by_xpath("//div[@class='paginator greyBox ']/ul/a[@class='paginator-button paginator-button-next']").click()
    except Exception:
        break;

driver.close()
