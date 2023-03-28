from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from TestEnvironment.GlobalClassMethods.MasterDataExcelReader import DataReadMaster
import time

class LoaderCls:
    @classmethod
    def LoaderMeth(cls,driver):
        #pass
        SHORT_TIMEOUT = 2
        LONG_TIMEOUT = 60
        LOADING_ELEMENT_XPATH = DataReadMaster.GlobalData("GlobalData", "Loader")
        try:
            WebDriverWait(driver, SHORT_TIMEOUT
                          ).until(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))

            WebDriverWait(driver, LONG_TIMEOUT
                          ).until_not(EC.presence_of_element_located((By.XPATH, LOADING_ELEMENT_XPATH)))
        except TimeoutException:
            pass
        time.sleep(1)