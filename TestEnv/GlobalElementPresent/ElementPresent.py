from selenium.webdriver.common.by import By
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster

class ElementPresentCls:
    @classmethod
    def ElementPresentMeth(cls,driver,MdataSheetTab,MdataSheetItem,ElementExpected,ElementVerify,PageName,TestResult,TestResultStatus):
        try:
            ElementFound = driver.find_element(By.XPATH,
                                                   DataReadMaster.GlobalData(MdataSheetTab,
                                                                             MdataSheetItem)).text
            assert ElementExpected in ElementFound, ElementVerify + " at " + PageName + " not found"
            TestResult.append(ElementVerify + " at " + PageName + " was present")
            TestResultStatus.append("Pass")
        except Exception as e1:
            print(e1)
            TestResult.append(ElementVerify + " at " + PageName + " not found")
            TestResultStatus.append("Fail")