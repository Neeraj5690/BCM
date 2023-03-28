import time

from selenium import webdriver
from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.common.by import By
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster
from TestEnvironment.GlobalErrorPresent.ErrorPresent import ErrorPresentCls
from TestEnvironment.GlobalLoader.Loader import LoaderCls
from chrome.LatestChrome import ChromeCls

class FormFillCls:
    #driver = webdriver.Chrome(executable_path=ChromeCls.NewChromePath1ChrCls)
    @classmethod
    def FormFillMeth(cls,driver,ItemList,MdataSheetTab,MdataSheetItem,MdataSheetItem2,ElementExpected,ElementVerify,PageName,TestResult,TestResultStatus):
        try:
            LoaderCls.LoaderMeth(driver)
            driver.find_element(By.XPATH,
                                DataReadMaster.GlobalData(MdataSheetTab,MdataSheetItem)).click()
            LoaderCls.LoaderMeth(driver)
            MdataSheetItem=MdataSheetItem2
            ElementFound = driver.find_element(By.XPATH,
                                               DataReadMaster.GlobalData(MdataSheetTab,
                                                                         MdataSheetItem)).text
            # print("ElementExpected is "+ElementExpected)
            # print("ElementFound is "+ElementFound)
            # ---------------loop for all fields----------
            for ii in range(len(ItemList)):
                time.sleep(1)
                FieldID = ItemList[ii]
                #print("FieldID in List: " + FieldID)
                FieldIDExcel=DataReadMaster.GlobalDataForm(MdataSheetTab,ItemList[ii])
                # print("FieldID 1 from Excel: " + str(FieldIDExcel[0]))
                # print("FieldID 2 from Excel: " + str(FieldIDExcel[1]))
                substr = "$"
                x = ItemList[ii].split(substr)
                if x[1]=="Str":
                    driver.find_element(By.XPATH,str(FieldIDExcel[0])).send_keys(str(FieldIDExcel[1]))
                elif x[1]=="Drp":
                    for ii1 in range(20):
                        time.sleep(1)
                        driver.find_element(By.XPATH, str(FieldIDExcel[0])).click()
                        time.sleep(1)
                        ActionChains(driver).key_down(Keys.DOWN).perform()
                        time.sleep(1)
                        ActionChains(driver).key_down(Keys.ENTER).perform()
                        time.sleep(1)
                        Val=driver.find_element(By.XPATH, str(FieldIDExcel[0])+"/span").text
                        if Val == str(FieldIDExcel[1]):
                            print("**** Value found "+Val)
                            break
                        else:
                            print("Value not found "+Val)

                elif x[1]=="Clk":
                    driver.find_element(By.XPATH,str(FieldIDExcel[0]+"["+str(FieldIDExcel[1])+"]")).click()
                else:
                    print(x[1])

            assert ElementExpected in ElementFound, ElementVerify + " at " + PageName + " not found"
            TestResult.append(ElementVerify + " at " + PageName + " was present and working as expected")
            TestResultStatus.append("Pass")
        except Exception as e1:
            print(e1)
            TestResult.append(ElementVerify + " at " + PageName + " was either not found or not working")
            TestResultStatus.append("Fail")
            ErrorPresentCls.ErrorPresentMeth(driver, PageName, TestResult, TestResultStatus)