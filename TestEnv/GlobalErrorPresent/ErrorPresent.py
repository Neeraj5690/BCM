import pytest
from selenium.webdriver.common.by import By
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster

class ErrorPresentCls:
    @classmethod
    def ErrorPresentMeth(cls,driver,PageName,TestResult,TestResultStatus):
        for x in range(1, 5):
            try:
                ErrorToFound = driver.find_element(By.XPATH,
                                                   DataReadMaster.GlobalData("GlobalData",
                                                                             "Error"+str(x))).text
                TestResult.append("Below error found \n"+ErrorToFound)
                TestResultStatus.append("Fail")
                break
            except Exception:
                pass


