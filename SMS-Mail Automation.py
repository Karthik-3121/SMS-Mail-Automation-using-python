from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from tkinter import *
from tkinter import filedialog
import time
import pandas


def Send_Sms():
    driver = webdriver.Chrome("****") #Enter the location of Chormedriver 
    df=pandas.read_excel(path)
    driver.maximize_window()  #maximize the window size  
    driver.delete_all_cookies()  #delete the cookies  
    driver.get("http://192.168.8.1")  #HUAWEI Dongle URL->>MAYBE Change with model
    time.sleep(3)
    driver.find_element_by_id("sms_gif").click()
    time.sleep(2)
    driver.find_element_by_id("username").send_keys("***")#Enter the username HUAWEI Dongle website
    time.sleep(1)
    driver.find_element_by_id("password").send_keys("***")#Enter the the password
    time.sleep(1)
    driver.find_element_by_id("pop_login").click()#login
    time.sleep(1)
    for i,j in df.iterrows():
        driver.find_element_by_id("message").click()
        time.sleep(1)
        driver.find_element_by_id("recipients_number").send_keys(j['number'])
        time.sleep(1)
        driver.find_element_by_id("message_content").send_keys(j['content'])
        time.sleep(1)
        driver.find_element_by_id("pop_send").click()
        time.sleep(5)
        driver.find_element_by_id("pop_OK").click()
        time.sleep(1)
    time.sleep(2)
    driver.find_element_by_id("logout_span").click()#logout
    time.sleep(1)
    driver.find_element_by_id("pop_confirm").click()
    time.sleep(2)
    driver.close()#close the broswer

def Send_Mail():
    driver = webdriver.Chrome("****") #Enter the location of Chormedriver 
    df=pandas.read_excel(path) 
    driver.maximize_window()  #maximize the window size  
    driver.delete_all_cookies()  #delete the cookies  
    driver.get("https://mail.google.com")  #Gmail Url
    time.sleep(3)  
    driver.find_element_by_id("identifierId").send_keys("****") #Enter the Mail ID
    driver.find_element_by_id("identifierNext").click()  
    time.sleep(5)   
    driver.find_element_by_name("password").send_keys("*****") #Enter the Mail Password
    driver.find_element_by_class_name("VfPpkd-vQzf8d").click()
    time.sleep(10) 
    for i,j in df.iterrows():
        driver.find_element_by_class_name("T-I.T-I-KE.L3").click()
        driver.find_element_by_tag_name("textarea").send_keys(j['email'])
        time.sleep(2)  
        driver.find_element_by_name("subjectbox").send_keys(j['subject'])
        driver.find_element_by_class_name("Am.Al.editable.LW-avf.tS-tW").send_keys(j['content'])
        driver.find_element_by_class_name("T-I.J-J5-Ji.aoO.v7.T-I-atl.L3").click()
        time.sleep(5)
    driver.find_element_by_class_name("gb_Ca.gbii").click()
    driver.find_element_by_id("gb_71").click()
    driver.close()

def browseFiles():
	filename = filedialog.askopenfilename(initialdir = "/",
										title = "Select a File",
										filetypes = (("Microsoft Excel Worksheet",
														"*.xlsx*"),
													("all files",
														"*.*")))
	
	# Change label contents
	label_file_explorer.configure(text="File Opened: "+filename)
	global path
	path=filename
	return filename
																					
# Create the root window
window = Tk()
window.title('Automation with Python')
window.geometry("500x500")
window.config(background = "white")
label_file_explorer = Label(window,
							text = "Select File",
							width = 100, height = 4,
							fg = "blue")
button_explore = Button(window,
						text = "Browse Files",
						command = browseFiles)
path=""
label_Options = Label(window,
							text = "Select your option",
							width = 100, height = 4,
							fg = "blue")
button_ch1 = Button(window,
					text = "Send SMS",
					command =Send_Sms)
button_ch2 = Button(window,
					text = "Send Mail",
					command = Send_Mail)
button_exit = Button(window,
					text = "Exit",
					command = exit)
label_file_explorer.grid(column = 1, row = 1)
button_explore.grid(column = 1, row = 2)
label_Options.grid(column = 1, row = 3)
button_ch1.grid(column = 1,row = 4)
button_ch2.grid(column = 1,row = 5)
button_exit.grid(column = 1,row = 6)
window.mainloop()

