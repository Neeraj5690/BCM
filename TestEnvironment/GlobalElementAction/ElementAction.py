from selenium.webdriver.common.by import By
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster
from TestEnvironment.GlobalErrorPresent.ErrorPresent import ErrorPresentCls
from TestEnvironment.GlobalLoader.Loader import LoaderCls

class ElementActionCls:
    @classmethod
    def ElementActionMeth(cls,driver,MdataSheetTab,MdataSheetItem,MdataSheetItem2,ElementExpected,ElementVerify,PageName,TestResult,TestResultStatus):
        try:
            LoaderCls.LoaderMeth(driver)
            driver.find_element(By.XPATH,
                                DataReadMaster.GlobalData(MdataSheetTab,MdataSheetItem)).click()
            LoaderCls.LoaderMeth(driver)
            ElementFound = driver.find_element(By.XPATH,
                                               DataReadMaster.GlobalData(MdataSheetTab,
                                                                         MdataSheetItem2)).text
            # print("ElementExpected is "+ElementExpected)
            # print("ElementFound is "+ElementFound)
            assert ElementExpected in ElementFound, ElementVerify + " at " + PageName + " not found"
            TestResult.append(ElementVerify + " at " + PageName + " was present and working as expected")
            TestResultStatus.append("Pass")
        except Exception as e1:
            TestResult.append(ElementVerify + " at " + PageName + " was either not found or not working")
            TestResultStatus.append("Fail")
            ErrorPresentCls.ErrorPresentMeth(driver, PageName, TestResult, TestResultStatus)