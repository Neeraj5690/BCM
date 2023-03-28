import time

import pytest
from selenium.webdriver.common.by import By
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster
from TestEnvironment.GlobalLoader.Loader import LoaderCls
SafeToVerify=None
class SafeToElementActionCls:
    @classmethod
    def SafeToElementActionMeth(cls, driver,SafeToVerify,MdataSheetTab, MdataSheetItem):
        try:
            IfElementFound = driver.find_element(By.XPATH,
                                               DataReadMaster.GlobalData(MdataSheetTab,
                                                                         MdataSheetItem)).text
            if "No items available" == IfElementFound:
                SafeToVerify = "No"
                return SafeToVerify
            elif "No" in IfElementFound:
                SafeToVerify = "No"
                return SafeToVerify
            else:
                SafeToVerify = "Yes"
                return SafeToVerify
        except Exception as e1:
            print(e1)
            SafeToVerify = "No ex"
            return SafeToVerify