""" 
//*******************************
Copyright (c) 2017 Microsoft Corporation. All rights reserved. 
//*******************************"""

import unittest
import os
import time
from appium import webdriver
from selenium.webdriver.common.keys import Keys

#get the request and response details, request life cycle, 
#the moment when you send messages over http there should be something 
#listening to the my request currently my local host listens to the request 

class WordPadTests(unittest.TestCase):

    def setUp(self):
        # This sets up the appium webdriver
        desired_caps = {}
        desired_caps["app"] = "C:\\Program Files\\Windows NT\\Accessories\\wordpad.exe"
        self.driver = webdriver.Remote(
            command_executor='http://127.0.0.1:4723',
            desired_capabilities= desired_caps)

    def tearDown(self):
        self.driver.find_element_by_class_name("WordPadClass").send_keys(Keys.ALT + Keys.F4)
        donotsave = self.driver.find_element_by_name("Don't Save")
        donotsave.click()
        self.driver.quit()

    def test_write(self):
        self.driver.find_element_by_class_name("RICHEDIT50W").send_keys("Hello World :)")
        self.driver.find_element_by_class_name("RICHEDIT50W").clear()



    def test_initialize(self):
        # This checks to see if word pad opened up with a new window
        #self.assertIsNotNull()
        titlebar = self.driver.find_element_by_accessibility_id("TitleBar")
        self.assertTrue(str(titlebar.text), "Document -WordPad")
        textwindow = self.driver.find_element_by_class_name("RICHEDIT50W")
        self.assertTrue(str(textwindow.text), "Rich Text Window")

    def test_zoom(self):
        self.driver.find_element_by_class_name("RICHEDIT50W").send_keys("This is not a drill")
        self.driver.find_element_by_name("View").click()
        self.driver.find_element_by_name("Zoom").click()
        self.driver.find_element_by_name("Zoom out").click()
        self.driver.find_element_by_name("100%").click()

'''
    def test_valid(self):
        self.assertTrue(str(filetab.text), "File tab")
        filetab.driver.click
        filebox = self.driver.find_element_by_class_name("NetUIHWND")
        self.assertTrue(str(filebox.text), "")
        self.driver.find_element_by_name("Save").click
        #finds the save as box
        saveasbox = self.driver.find_element_by_class_name("#32770")
        self.assertTrue(str(saveasbox.text), "Save As")
        #finds the file name box in save as
        filenamebox = self.driver.find_element_by_accessibility_id("1001")
        self.assertTrue(str(filenamebox.text), "File name")
        filenamebox = self.driver.find_element_by_accessibility_id("1001")
        filenamebox.driver.clear
        filenamebox.driver.send_keys("This is my test")
        # Save Button by automation ID
        self.driver.find_element_by_accessibility_id("1").click
'''    
     
     
if __name__ == '__main__':
        suite = unittest.TestLoader().loadTestsFromTestCase(WordPadTests)
        unittest.TextTestRunner(verbosity=2).run(suite)




