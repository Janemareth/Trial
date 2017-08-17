import unittest
import os
from appium import webdriver

class outlooktest(unittest.TestCase):
    
    @classmethod
    def setUpClass(cls):
        # This sets up the appium webdriver
        desired_caps = {}
        desired_caps["app"] = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\outlook.exe"
        cls.driver = webdriver.Remote(
            command_executor='http://127.0.0.1:4723',
            desired_capabilities= desired_caps)
        #if test fails to open
        #raise Exception ("Unable to open Outlook")

    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()

    def test_other_focused(self):
        titlebar = self.driver.find_element_by_class_name("NetUIOfficeCaption")
        self.assertTrue(str(titlebar.text), "Inbox - jamareth@microsoft.com - Outlook") 
        #need to find a way to dynamically change the titlebar string so the email is skipped over or changes automatically
        #perhaps it could be a get operation instead
        self.driver.find_element_by_name("Other").click
        self.driver.find_element_by_name("Focused").click

    def test_save(self):

        self.driver.find_element_by_name("View").click
        #self.driver.find_element_by_name("Arrange by").click
        self.driver.find_element_by_name("From").click

    def test_help (self):

        self.driver.find_element_by_name("Help").click
        helpsupport = self.driver.find_element_by_class_name("NetUIChunk")
        self.assertTrue(str(helpsupport.text), "Help & Support")
        self.driver.find_element_by_name("Show Training").click
        self.driver.find_element_by_class_name("MsoCommandBar")
        elem = self.driver.find_element_by_accessibility_id("ocSearchBox")
        elem.driver.send_keys("How to send email")
        self.driver.find_element_by_accessibility_id("ocSearchButton").click


if __name__ == '__main__':
    suite = unittest.TestLoader().loadTestsFromTestCase(outlooktest)
    unittest.TextTestRunner(verbosity=2).run(suite)
