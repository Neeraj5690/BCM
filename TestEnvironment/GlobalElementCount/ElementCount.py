import time
from selenium.webdriver.common.by import By
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster
from TestEnvironment.GlobalLoader.Loader import LoaderCls

class ElementCountCls:
    @classmethod
    def ElementCountMeth(cls,driver,MdataSheetTab,MdataSheetItem,ElementExpected,ElementVerify,PageName,TestResult,TestResultStatus):
        try:
            ItemsPerPage = driver.find_elements(By.XPATH,
                                                   DataReadMaster.GlobalData(MdataSheetTab,
                                                                             MdataSheetItem))
            ItemsPerPage=len(ItemsPerPage)
            try:
                while driver.find_element(By.XPATH,
                    "//div[@class='GridFooter---align_end']/span[4]/a[2][@title='Last page']").is_displayed() == True:
                    time.sleep(1)
                    LoaderCls.LoaderMeth(driver)
                    driver.find_element(By.XPATH,
                                        "//div[@class='GridFooter---align_end']/span[4]/a[1][@title='Next page']").click()
                    LoaderCls.LoaderMeth(driver)
                    ItemsNextPage = driver.find_elements(By.XPATH,
                                                        DataReadMaster.GlobalData(MdataSheetTab,
                                                                                  MdataSheetItem))
                    ItemsPerPage=ItemsPerPage+len(ItemsNextPage)
            except Exception:
                print("No last page footer icon found")
                pass
            ElementFound=str(ItemsPerPage)
            assert ElementExpected in ElementFound, ElementVerify + " count [ "+ElementExpected+" ] at " + PageName + " was not correct. Found [ "+ElementFound+" ]"
            TestResult.append(ElementVerify + " count [ "+ElementFound+" ] at " + PageName + " was correct")
            TestResultStatus.append("Pass")
        except Exception as e1:
            print(e1)
            TestResult.append(ElementVerify + " count [ "+ElementExpected+" ] at " + PageName + " was not correct. Found [ "+ElementFound+" ]")
            TestResultStatus.append("Fail")