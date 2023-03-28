import time

import pytest
from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.common.by import By
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster
from TestEnvironment.GlobalErrorPresent.ErrorPresent import ErrorPresentCls
from TestEnvironment.GlobalLoader.Loader import LoaderCls

class ElementActionClsNewTab:
    @classmethod
    def ElementActionMethNewTab(cls,driver,MdataSheetTab,MdataSheetItem,MdataSheetItem2,ElementExpected,ElementVerify,PageName,TestResult,TestResultStatus):
        try:
            driver.find_element(By.XPATH,
                                DataReadMaster.GlobalData(MdataSheetTab,MdataSheetItem)).click()
            time.sleep(1)
            LoaderCls.LoaderMeth(driver)
            driver.switch_to.window(driver.window_handles[1])
            MdataSheetItem = MdataSheetItem2
            ElementFound = driver.find_element(By.XPATH,
                                               DataReadMaster.GlobalData(MdataSheetTab,
                                                                         MdataSheetItem)).text
            print("&&&&&&&&&&&&&&&&&&& ElementFound is "+ElementFound)
            for winclose in range(1, 10):
                time.sleep(1)
                if len(driver.window_handles) > 1:
                    driver.switch_to.window(driver.window_handles[1])
                    driver.close()
                elif len(driver.window_handles) == 1:
                    break
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(1)
            assert ElementExpected in ElementFound, ElementVerify + " at " + PageName + " not found"
            TestResult.append(ElementVerify + " at " + PageName + " was present and working as expected")
            TestResultStatus.append("Pass")
        except Exception as e1:
            TestResult.append(ElementVerify + " at " + PageName + " was either not found or not working")
            TestResultStatus.append("Fail")
            ErrorPresentCls.ErrorPresentMeth(driver, PageName, TestResult, TestResultStatus)