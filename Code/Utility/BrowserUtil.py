from robot.libraries.BuiltIn import BuiltIn
from Selenium2Library.keywords.keywordgroup import KeywordGroup
from selenium.common.exceptions import NoAlertPresentException 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from selenium.common.exceptions import StaleElementReferenceException
from robot.api import logger
from Selenium2Library.locators import ElementFinder


class BrowserUtil (KeywordGroup):

    def is_alert_exist(self):
        try:
            s2Lib = BuiltIn().get_library_instance('Selenium2Library')
            browser = s2Lib._current_browser()
            alert = browser.switch_to_alert()
            # if Alert does not exist, then alert.text will cause the exception.
            text = alert.text
            return True
        except    NoAlertPresentException, e:
            return "False - No Alert"
        # Handle rest of exceptions including no browser.
        except:
            return "False - No Browser"
            
    def TypeToAlertField(self,textToType):
        try:
            s2Lib = BuiltIn().get_library_instance('Selenium2Library')
            browser = s2Lib._current_browser()
            alert = browser.switch_to_alert()          
            alert.send_keys(textToType)
            alert.send_keys(Keys.TAB)
            return True
        except    NoAlertPresentException, e:
            return "Either No Alert Or Field Missing"      
        except:
            return "Writing To Field Missing"

    def get_CSS_Property_Value(self,locator,CSSproperty):
        s2l = BuiltIn().get_library_instance('Selenium2Library')
        element = s2l._element_find(locator, True, False)
        propValue = element.value_of_css_property(CSSproperty)
        return propValue

    def switch_to_second_window(self):
        self.s21 = BuiltIn().get_library_instance('Selenium2Library')
        browser = self.s21._current_browser()
        for handle in browser.window_handles:
            print(handle)
            lstHandle = handle
        browser.switch_to_window(lstHandle)
        new = browser.title
        print(new)   