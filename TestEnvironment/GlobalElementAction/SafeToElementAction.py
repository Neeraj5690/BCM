import math
import re
import time

from selenium.webdriver.common.by import By
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster
from TestEnvironment.GlobalErrorPresent.ErrorPresent import ErrorPresentCls
from TestEnvironment.GlobalLoader.Loader import LoaderCls

SafeToVerify=None
class SafeToElementActionCls:
    @classmethod
    def SafeToElementTable(cls, driver,MdataSheetTab, MdataSheetItem):
        try:
            IfElementFound = driver.find_element(By.XPATH,DataReadMaster.GlobalData(MdataSheetTab,MdataSheetItem)).text
            if "no" in IfElementFound or "No" in IfElementFound:
                SafeToVerify = "No"
                return SafeToVerify
            else:
                SafeToVerify = "Yes"
                return SafeToVerify
        except Exception as e1:
            print(e1)
            SafeToVerify = "Safe to Table element exception found"
            return SafeToVerify

    @classmethod
    def SafeToTableFooterNav(cls, driver,ElementVerify, MdataSheetTab, MdataSheetItem,EleCount,NextClick,PageName,TestResult,TestResultStatus):
        try:
            ElementFound = driver.find_element(By.XPATH,
                                                 DataReadMaster.GlobalData(MdataSheetTab, MdataSheetItem)).text
            print("TotalItem " + ElementFound)
            substr = "of"
            if substr in ElementFound:
                x = ElementFound.split(substr)
                string_name = x[0]
                TotalItemAfterOf = x[1]
                print("string_name "+string_name)
                print("TotalItemAfterOf "+TotalItemAfterOf)
                substr = "–"
                try:
                    if substr in string_name:
                        x = string_name.split(substr)
                        string_nameBefore = x[0]
                        string_nameAfter = x[1]
                        print("string_nameBefore " + string_nameBefore)
                        print("string_nameAfter " + string_nameAfter)
                        IterateNo = int(TotalItemAfterOf) / int(string_nameAfter)
                        print(str(float(IterateNo)))
                        IterateNo = math.ceil(float(IterateNo))
                        print(IterateNo)

                        ElementCount=driver.find_elements(By.XPATH,DataReadMaster.GlobalData(MdataSheetTab, EleCount))
                        ElementCount=len(ElementCount)
                        for NextClickCount in range(1, IterateNo):
                            print("NextClickCount "+str(NextClickCount))
                            time.sleep(1)
                            driver.find_element(By.XPATH,DataReadMaster.GlobalData(MdataSheetTab, NextClick)).click()
                            time.sleep(1)
                            ElementCountNext = driver.find_elements(By.XPATH,
                                                                DataReadMaster.GlobalData(MdataSheetTab, EleCount))
                            ElementCountNext = len(ElementCountNext)
                            ElementCount=ElementCount+ElementCountNext

                        print("ElementCountFound "+str(ElementCount))
                        if ElementCount==int(TotalItemAfterOf):
                            TestResult.append(ElementVerify +" at "+PageName+ " was working as expected")
                            TestResultStatus.append("Pass")
                        else:
                            TestResult.append(ElementVerify+" at "+PageName+ " wasn't working as expected")
                            TestResultStatus.append("Fail")
                            ErrorPresentCls.ErrorPresentMeth(driver, PageName, TestResult, TestResultStatus)
                except Exception as qq:
                    print(qq)
        except Exception as e1:
            print(e1)
            SafeToVerify = "Safe to Table element exception found- "+e1
            return SafeToVerify