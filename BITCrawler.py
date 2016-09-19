#-*- coding:utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import xlwt
import traceback
import socket

url = "http://www.bit.edu.cn"
username = '1120101866'
passwd = '411425199102146631'
print url

browser = webdriver.Ie()
# 登录
print 'step1'
loginUrl = 'http://10.5.2.80/(l2ay3p55k13s5m45zkn1xy55)/default2.aspx'
try:
	browser.get(loginUrl)
	time.sleep(1)
	browser.save_screenshot('step1.png')

	print 'step2'
	browser.find_element_by_name('TextBox1').send_keys(username)
	browser.find_element_by_name('TextBox1').send_keys(Keys.TAB)
	browser.find_element_by_name('TextBox2').send_keys(passwd)
	browser.find_element_by_name('TextBox2').send_keys(Keys.ENTER)
	# 等待加载成功
	locator = (By.CLASS_NAME, "MainMenu")
	WebDriverWait(browser, 50).until(EC.presence_of_element_located(locator))
	#time.sleep(5)
	browser.save_screenshot('step2.png')

	print 'step3'
	# 获取成绩
	queryElement = browser.find_elements_by_class_name("MainMenu")[3]
	ActionChains(browser).move_to_element(queryElement).perform()
	browser.save_screenshot('step3.png')

	# 
	print 'step4'
	scoreElement = browser.find_element_by_id("SubMenuN1216").find_elements_by_class_name("SubMenu")[2]
	# 保留当前窗口
	nowHandle = browser.current_window_handle
	scoreElement.click()
	# 所有窗口
	allHandles = browser.window_handles
	for handle in allHandles:
		if handle != nowHandle: #跳转到新的窗口
			browser.switch_to_window(handle)
	browser.save_screenshot('step4.png')
	print browser.current_url

	print 'step5'
	# 点击获取所有成绩
	allScoresBtn = browser.find_element_by_name("Button2")
	allScoresBtn.click()
	time.sleep(5)
	browser.save_screenshot('step5.png')

	# 等待加载成功
	locator = (By.CLASS_NAME, "datagridstyle")
	WebDriverWait(browser, 50).until(EC.presence_of_element_located(locator))
	scoreTable = browser.find_element_by_class_name('datagridstyle')
	# row = scoreTable.find_element_by_xpath("//tr[1]")
	# print row.text
	# row = row.find_element_by_xpath('//following-sibling::tr[1]')
	# print row.text
	rows = scoreTable.find_elements_by_tag_name('tr')


	print 'step6'
	# 保存数据到excel中
	filename = username+'.xls'
	f = xlwt.Workbook()
	sheet1 = f.add_sheet(u'score', cell_overwrite_ok=True)
	for rowindex, row in enumerate(rows):
		cols = row.find_elements_by_tag_name('td')
		for colindex, col in enumerate(cols):
			text = col.text
			if text:
				text = text.strip().replace('&nbsp;', '')

			sheet1.write(rowindex, colindex, text)

	f.save(filename)


except Exception as e:
	print traceback.format_exc()
finally:
	#browser.close()
	browser.quit()
	print 'done'
