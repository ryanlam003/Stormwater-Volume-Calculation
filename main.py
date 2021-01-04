from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time


def extract_columns_as_arrays():
    # load the excel spreadsheet with all the values
    wb = load_workbook('Stormwater Credit List.xlsx')
    sheet = wb['Master']

    # initialize a list of addresses, total areas, percent pervious area, disconnection area, green roof area,...
    # permeable pavement area, infiltration basin area, rain gardens area, rain harvesting area, and impervious area
    addressList = []
    totalAreaList = []
    imperviousAreaList = []
    percentPerviousList = []
    disconnectionAreaList = []
    greenRoofAreaList = []
    permeablePavementAreaList = []
    infiltrationBasinAreaList = []
    rainGardensAreaList = []
    rainHarvestingAreaList = []

    # populate the address list, adding 'Philadelphia'
    for columnOfCellObjects in sheet['D3':'D79']:
        for cellObj in columnOfCellObjects:
            addressList.append(cellObj.value)
    for ii in range(0, len(addressList)):
        addressList[ii] = addressList[ii] + ' Philadelphia'

    # populate the total area list, converting the areas from square feet to acres
    for columnOfCellObjects in sheet['H3':'H79']:
        for cellObj in columnOfCellObjects:
            totalAreaList.append(cellObj.value)
    totalAreaListAcres = totalAreaList
    for ii in range(0, len(totalAreaListAcres)):
        totalAreaListAcres[ii] = totalAreaListAcres[ii] / 43560

    # populate the impervious area list
    for columnOfCellObjects in sheet['I3':'I79']:
        for cellObj in columnOfCellObjects:
            imperviousAreaList.append(cellObj.value)

    # populate the percent pervious area list
    for columnOfCellObjects in sheet['K3':'K79']:
        for cellObj in columnOfCellObjects:
            percentPerviousList.append(cellObj.value)

    # populate the disconnection area list, convert to percentages
    for columnOfCellObjects in sheet['O3':'O79']:
        for cellObj in columnOfCellObjects:
            disconnectionAreaList.append(cellObj.value)
    percentDisconnection = disconnectionAreaList
    for ii in range(0, len(percentDisconnection)):
        percentDisconnection[ii] = 100 * percentDisconnection[ii] / imperviousAreaList[ii]

    # populate the green roof area list, convert to percentages
    for columnOfCellObjects in sheet['P3':'P79']:
        for cellObj in columnOfCellObjects:
            greenRoofAreaList.append(cellObj.value)
    percentGreenRoof = greenRoofAreaList
    for ii in range(0, len(percentGreenRoof)):
        percentGreenRoof[ii] = 100 * percentGreenRoof[ii] / imperviousAreaList[ii]

    # populate the permeable pavement area list, convert to percentages
    for columnOfCellObjects in sheet['Q3':'Q79']:
        for cellObj in columnOfCellObjects:
            permeablePavementAreaList.append(cellObj.value)
    percentPermeablePavement = permeablePavementAreaList
    for ii in range(0, len(percentPermeablePavement)):
        percentPermeablePavement[ii] = 100 * percentPermeablePavement[ii] / imperviousAreaList[ii]

    # populate the infiltration basin area list, convert to percentages
    for columnOfCellObjects in sheet['U3':'U79']:
        for cellObj in columnOfCellObjects:
            infiltrationBasinAreaList.append(cellObj.value)
    percentInfiltration = infiltrationBasinAreaList
    for ii in range(0, len(percentInfiltration)):
        percentInfiltration[ii] = 100 * percentInfiltration[ii] / imperviousAreaList[ii]

    # populate the rain gardens area list, convert to percentages
    for columnOfCellObjects in sheet['V3':'V79']:
        for cellObj in columnOfCellObjects:
            rainGardensAreaList.append(cellObj.value)
    percentRainGarden = rainGardensAreaList
    for ii in range(0, len(percentRainGarden)):
        percentRainGarden[ii] = 100 * percentRainGarden[ii] / imperviousAreaList[ii]

    # populate the rain harvesting area list, convert to percentages
    for columnOfCellObjects in sheet['T3':'T79']:
        for cellObj in columnOfCellObjects:
            rainHarvestingAreaList.append(cellObj.value)
    percentRainHarvesting = rainHarvestingAreaList
    for ii in range(0, len(percentRainHarvesting)):
        percentRainHarvesting[ii] = 100 * percentRainHarvesting[ii] / imperviousAreaList[ii]

    return (addressList, totalAreaListAcres, percentPerviousList, percentDisconnection, percentGreenRoof,
            percentPermeablePavement, percentInfiltration, percentRainGarden, percentRainHarvesting)


(addressList, totalAreaListAcres, percentPerviousList, percentDisconnection, percentGreenRoof, percentPermeablePavement,
 percentInfiltration, percentRainGarden, percentRainHarvesting) = extract_columns_as_arrays()

for ii in range(0,len(addressList)):
    # using chrome to access web
    driver = webdriver.Chrome()

    # open the website
    driver.get('https://swcweb.epa.gov/stormwatercalculator/')
    time.sleep(1)
    siteNameBox = driver.find_element_by_xpath('//*[@id="homePage"]/div/div/div[1]/div/div/div/input')
    siteNameBox.send_keys(addressList[ii])
    getStartedButton = driver.find_element_by_xpath('//*[@id="getStartedButton"]')
    getStartedButton.click()
    time.sleep(2)

    # enter in the first address and area
    addressBox = driver.find_element_by_xpath('//*[@id="locationInput"]')
    addressBox.send_keys(addressList[ii])
    addressBox.send_keys(Keys.ENTER)
    time.sleep(2)
    siteAcreageBox = driver.find_element_by_xpath('//*[@id="acreInput"]')
    siteAcreageBox.clear()
    siteAcreageBox.send_keys(str(totalAreaListAcres[ii]))

    # click the soil type tab, no survey, no expansion
    soilTypeTab = driver.find_element_by_xpath('//*[@id="soiltype"]')
    soilTypeTab.click()
    time.sleep(6)
    surveyButtonNo = driver.find_element_by_xpath('//*[@id="fsrInvite"]/section[3]/button[2]')
    surveyButtonNo.click()
    time.sleep(4)
    expandSearchButtonNo = driver.find_element_by_xpath('//*[@id="expandSearchButtons"]/div[2]/button')
    expandSearchButtonNo.click()
    time.sleep(1)

    # choose sandy loam
    sandyLoamRadio = driver.find_element_by_xpath('//*[@id="sandyLoamRadio"]')
    sandyLoamRadio.click()

    # choose flat topography
    topographyTab = driver.find_element_by_xpath('//*[@id="topography"]')
    topographyTab.click()
    time.sleep(2)
    flatRadio = driver.find_element_by_xpath('//*[@id="flatRadio"]')
    flatRadio.click()

    # set land cover percentages
    landCoverTab = driver.find_element_by_xpath('//*[@id="land"]')
    landCoverTab.click()
    time.sleep(2)
    lawnPercentBox = driver.find_element_by_xpath('//*[@id="lawnValue"]')
    lawnPercentBox.clear()
    lawnPercentBox.send_keys(str(round(percentPerviousList[ii])))

    # set Low Impact Development (LID)/Green Stormwater Management Infrastructure controls
    LIDcontrolsTab = driver.find_element_by_xpath('//*[@id="lid"]')
    LIDcontrolsTab.click()
    time.sleep(2)
    disconnectionBox = driver.find_element_by_xpath('//*[@id="disconnectionValue"]')
    disconnectionBox.clear()
    disconnectionBox.send_keys(str(round(percentDisconnection[ii])))
    rainHarvestingBox = driver.find_element_by_xpath('//*[@id="rainHarvestingValue"]')
    rainHarvestingBox.clear()
    rainHarvestingBox.send_keys(str(round(percentRainHarvesting[ii])))
    rainGardenBox = driver.find_element_by_xpath('//*[@id="rainGardensValue"]')
    rainGardenBox.clear()
    rainGardenBox.send_keys(str(round(percentRainGarden[ii])))
    greenRoofBox = driver.find_element_by_xpath('//*[@id="greenRoofsValue"]')
    greenRoofBox.clear()
    greenRoofBox.send_keys(str(round(percentGreenRoof[ii])))
    infiltrationBasinBox = driver.find_element_by_xpath('//*[@id="infiltrationBasinsValue"]')
    infiltrationBasinBox.clear()
    infiltrationBasinBox.send_keys(str(round(percentInfiltration[ii])))
    permeablePavementBox = driver.find_element_by_xpath('//*[@id="permeablePavementValue"]')
    permeablePavementBox.clear()
    permeablePavementBox.send_keys(str(round(percentPermeablePavement[ii])))

    stormSizeBox = driver.find_element_by_xpath('//*[@id="designStormValue"]')
    stormSizeBox.clear()
    stormSizeBox.send_keys('2')

    # get results
    resultsTab = driver.find_element_by_xpath('//*[@id="results"]')
    driver.execute_script("arguments[0].click();", resultsTab)
    time.sleep(2)
    resultsTab.click()
    time.sleep(2)
    refreshResultsButton = driver.find_element_by_xpath('//*[@id="refreshResultsButton"]')
    refreshResultsButton.click()
    time.sleep(70)
    printResultsButton = driver.find_element_by_xpath('//*[@id="resultsPDFButton"]')
    printResultsButton.click()
    time.sleep(5)
    driver.close()