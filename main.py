import requests
import csv
import xlsxwriter
import json
import os
import time
import math
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET
import pprint
from selenium.webdriver.firefox.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

import json
import xlrd
import os
import time
OUTPUT_DIR	= 'output'

def try_for(function=lambda: None, iterations=5, retry_message="retry +%tries%"):
	'''run the `function` for the set number of `iterations` until the `function` raises Exception\n
	Arguments:
	`function`: the function to run\n
	`iterations`: number of times to run the `function` until it runs successfully without any Exception\n
	`retry_message`: message to display while retrying to run the `function`

	Returns:
		`ouput`, `error`
	'''

	error = None
	output = None

	tries = 0
	success = False

	while not success and tries < iterations:
		try:
			output = function()
			success = True
			print('MenuItems getting success!' )
		except Exception as e:
			success = False
			error = e
			print(str(tries + 1)+' %tries%' )

		tries += 1
		time.sleep(5)

	if not success and tries > iterations:
		error = Error("number of tries exceeded")

	return output, error


def get(url, headers=dict()) -> requests.Response:
	return requests.get(url, headers=headers)

def post(url, headers=dict()) -> requests.Response:
	return requests.post(url,data={},timeout=2.50)

def menuItems(url, headers=dict()):
	response = get(url, headers)
	if response.status_code != 200:
		raise Exception(f"Response ends with status code ({response.status_code})")

	return json.loads(response.text)

def read_categories():
	CAT_FILE="productUrl.xlsx"

	SHEET_NAME 		= 'Sheet1'
	URL_COL_NAME 	= 'url'
	CAT_COL_NAME	= 'catName'

	if not os.path.exists(CAT_FILE):
		log(f'{F_RED}Category file path does not exist{RESET_STYLE}')
		return None

	try:
		dfs = pd.read_excel(CAT_FILE, sheet_name=SHEET_NAME, header=0) # header row starts from 1
		cat_names 	= list(dfs[CAT_COL_NAME])
		urls		= list(dfs[URL_COL_NAME])

		if len(urls) != len(cat_names):
			return None

		categories 	= list()
		for index in range(len(cat_names)):
			container = dict()
			container["url"]=urls[index]
			container["catName"]=cat_names[index]
			categories=categories+[container]
		return categories
	except Exception:
		return None

if __name__ == '__main__':

	HEADERS = {
			'Content-Type': 'application/json',
			'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:84.0) Gecko/20100101 Firefox/84.0',
			'Origin': 'https://eshop.elit.sk',
			'Host': 'eshop.elit.sk',
			'Referer': 'https://eshop.elit.sk/Catalog/zimna-sezonna-ponuka/51314387',
			'x-cms': 'v2',
			'x-content': 'desktop',
			'x-mp': 'eshop.elit.sk',
			'x-platform': 'web',
			'Accept': 'application/json, text/plain, */*',
			'Connection': 'keep-alive',
			'Cache-Control': 'no-cache, max-age=0, must-revalidate, no-store',
			'Accept-Encoding': 'gzip, deflate, br',
			'TE': 'Trailers'
		}
	options = Options()
	options.add_argument('--headless')
    # driver = webdriver.Firefox(options=options)
	dt = datetime.now().strftime("%d%m%Y%H%M%S")
	menuItems, error  = try_for(lambda: menuItems('https://eshop.elit.sk/Catalog/MenuItems', HEADERS), 
				retry_message=f' [menuItems getting ] Retry + %tries%')
	# print(menuItems)
	data=''
	productData=[]
	products = ET.Element('products')
	if not os.path.exists("productUrl.xlsx"):
		# os.makedirs(OUTPUT_DIR)
		for menuItem in menuItems:
			categoryId=menuItem['CatalogId']
			categoryName=menuItem['Text']
			categoryLink=menuItem['Link']	

			driver = webdriver.Firefox(options=options)
			driver.implicitly_wait(2)
			driver.get("https://eshop.elit.sk"+categoryLink)
			items= driver.find_elements_by_class_name("CatalogRow-cat")
			if len(items) >0 :
				for item in items:
					catName=categoryName
					catName += "/"+ item.find_element_by_tag_name("a").text
					subUrl=item.find_element_by_tag_name("a").get_attribute("href")
					# print(catName,"============",subUrl)
					# driver1 = webdriver.Firefox(options=options)
					driver1 = webdriver.Firefox(options=options)
					driver1.implicitly_wait(2)
					driver1.get(subUrl)
					subItems=driver1.find_elements_by_class_name("CatalogRowOnly")
					if len(subItems) >0 :
						for subItem in subItems:
							subCatName=subItem.find_element_by_tag_name("a").text
							productsUrl=subItem.find_element_by_tag_name("a").get_attribute("href")
							productsCount=subItem.find_element_by_tag_name("span").text

							drive2 =webdriver.Firefox(options=options)
							drive2.implicitly_wait(2)
							drive2.get(productsUrl)
							subItems1=drive2.find_elements_by_class_name("CatalogRowOnly")
							if len(subItems1)>0:
								for subItem1 in subItems1:
									subCatName1=subItem1.find_element_by_tag_name("a").text
									productsUrl1=subItem1.find_element_by_tag_name("a").get_attribute("href")
									drive3 =webdriver.Firefox(options=options)
									drive3.implicitly_wait(2)
									drive3.get(productsUrl1)
									subItems2=drive3.find_elements_by_class_name("CatalogRowOnly")
									if len(subItems2)>0:
										for subItem2 in subItems2:
											subCatName2=subItem2.find_element_by_tag_name("a").text
											productsUrl2=subItem2.find_element_by_tag_name("a").get_attribute("href")
											drive4 =webdriver.Firefox(options=options)
											drive4.implicitly_wait(2)
											drive4.get(productsUrl2)
											subItmes3=drive4.find_elements_by_class_name("CatalogRowOnly")
											if len(subItmes3) > 0:
												for subItem3 in subItmes3:
													subCatName3=subItem3.find_element_by_tag_name("a").text
													productsUrl3=subItem3.find_element_by_tag_name("a").get_attribute("href")
													drive5 =webdriver.Firefox(options=options)
													drive5.implicitly_wait(2)
													drive5.get(productsUrl3)
													subItmes4=drive5.find_elements_by_class_name("CatalogRowOnly")
													if len(subItmes4) > 0:
														for subItem4 in subItmes4:
															subCatName4=subItem4.find_element_by_tag_name("a").text
															productsUrl4=subItem4.find_element_by_tag_name("a").get_attribute("href")
															drive6 =webdriver.Firefox(options=options)
															drive6.implicitly_wait(2)
															drive6.get(productsUrl4)
															subItmes5=drive6.find_elements_by_class_name("CatalogRowOnly")
															if len(subItmes5) > 0:
																for subItem5 in subItmes5:
																	subCatName5=subItem5.find_element_by_tag_name("a").text
																	productsUrl5=subItem5.find_element_by_tag_name("a").get_attribute("href")
																	drive7 =webdriver.Firefox(options=options)
																	drive7.implicitly_wait(2)
																	drive7.get(productsUrl5)
																	subItmes6=drive7.find_elements_by_class_name("CatalogRowOnly")
																	if len(subItmes6) > 0:
																		for subItem6 in subItmes6:
																			subCatName6=subItem6.find_element_by_tag_name("a").text
																			productsUrl6=subItem6.find_element_by_tag_name("a").get_attribute("href")
																			drive8 =webdriver.Firefox(options=options)
																			drive8.implicitly_wait(2)
																			drive8.get(productsUrl6)
																			subItmes7=drive8.find_elements_by_class_name("CatalogRowOnly")
																			if len(subItmes7) > 0:
																				for subItem7 in subItmes7:
																					subCatName7=subItem7.find_element_by_tag_name("a").text
																					productsUrl7=subItem7.find_element_by_tag_name("a").get_attribute("href")
																					drive9 =webdriver.Firefox(options=options)
																					drive9.implicitly_wait(2)
																					drive9.get(productsUrl6)
																					subItmes8=drive9.find_elements_by_class_name("CatalogRowOnly")
																					if len(subItmes8) > 0:
																						for subItem8 in subItmes8:
																							subCatName8=subItem8.find_element_by_tag_name("a").text
																							productsUrl7=subItem7.find_element_by_tag_name("a").get_attribute("href")
																							# drive10 =webdriver.Firefox(options=options)
																							# drive10.implicitly_wait(5)
																							# drive10.get(productsUrl6)
																							# subItmes9=driver10.find_elements_by_class_name("CatalogRowOnly")
																					else:
																						container = dict()
																						container["url"]=productsUrl7
																						container["catName"]=catName+"/"+subCatName+"/"+subCatName1+"/"+subCatName2 +"/"+subCatName3+"/"+subCatName4+"/"+subCatName5+"/"+subCatName6+"/"+subCatName7
																						productData=productData+[container]
																						print(productsUrl7 + " puted into Url list!")
																					drive9.quit()

																			else:
																				container = dict()
																				container["url"]=productsUrl6
																				container["catName"]=catName+"/"+subCatName+"/"+subCatName1+"/"+subCatName2 +"/"+subCatName3+"/"+subCatName4+"/"+subCatName5+"/"+subCatName6
																				productData=productData+[container]
																				print(productsUrl6 + " puted into Url list!")
																			drive8.quit()
																	else:
																		container = dict()
																		container["url"]=productsUrl5
																		container["catName"]=catName+"/"+subCatName+"/"+subCatName1+"/"+subCatName2 +"/"+subCatName3+"/"+subCatName4+"/"+subCatName5
																		productData=productData+[container]
																		print(productsUrl5 + " puted into Url list!")
																	drive7.quit()
																	try:
																		drive8.quit()
																	except:
																		print("")

															else:
																container = dict()
																container["url"]=productsUrl4
																container["catName"]=catName+"/"+subCatName+"/"+subCatName1+"/"+subCatName2 +"/"+subCatName3+"/"+subCatName4
																productData=productData+[container]
																print(productsUrl4 + " puted into Url list!")
															drive6.quit()
															try:
																drive7.quit()
															except:
																print("")


													else:
														container = dict()
														container["url"]=productsUrl3
														container["catName"]=catName+"/"+subCatName+"/"+subCatName1+"/"+subCatName2 +"/"+subCatName3
														productData=productData+[container]	
														print(productsUrl3 + " puted into Url list!")
													drive5.quit()
													try:
														drive6.quit()
													except:
														print("")


											else:
												container = dict()
												container["url"]=productsUrl2
												container["catName"]=catName+"/"+subCatName+"/"+subCatName1+"/"+subCatName2
												productData=productData+[container]	
												print(productsUrl2 + " puted into Url list!")
											drive4.quit()
											try:
												drive5.quit()
											except:
												print("")


									else:
										container = dict()
										container["url"]=productsUrl1
										container["catName"]=catName+"/"+subCatName+"/"+subCatName1
										productData=productData+[container]	
										print(productsUrl1 + " puted into Url list!")
									drive3.quit()
									try:
										drive4.quit()
									except:
										print("")

							else:
								container = dict()
								container["url"]=productsUrl
								container["catName"]=catName+"/"+subCatName
								productData=productData+[container]
								print(productsUrl + " puted into Url list!")
							drive2.quit()
							try:
								drive3.quit()
							except:
								print("")



					else:
						container = dict()
						container["url"]=subUrl
						container["catName"]=catName
						productData=productData+[container]
						print(subUrl + " puted into Url list!")

					driver1.quit()
					try:
						driver1.quit()
					except:
						print("")
					try:
						drive2.quit()
					except:
						print("")
					
					try:
						drive3.quit()
					except:
						print("")
					try:
						drive4.quit()
					except:
						print("")
					try:
						drive5.quit()
					except:
						print("")
					try:
						drive6.quit()
					except:
						print("")
					try:
						drive7.quit()
					except:
						print("")
					
					############break###################
					# break
				
			driver.quit()
			print(productData)

	
	

		# create output file name based on time
		dt = datetime.now().strftime("%d%m%Y%H%M%S")
		outputXLSX = os.path.join(f"productUrl.xlsx")

		workbook = xlsxwriter.Workbook(outputXLSX)
		worksheet = workbook.add_worksheet()
			
		BASIC_FORMAT = workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter'})

		worksheet.write("A1", "url", BASIC_FORMAT)
		worksheet.write("B1", "catName", BASIC_FORMAT)

		CELL_WIDTH = 150
		worksheet.set_column(0, CELL_WIDTH)
		row = 2
		print("writting...")
		for pd in productData:
			# print(em["email"])
			worksheet.write(f"A{row}", pd["url"], BASIC_FORMAT)
			worksheet.write(f"B{row}", pd["catName"], BASIC_FORMAT)
			row += 1

		workbook.close()
		print(f'+ data written to {outputXLSX}')
	else:
		productData=read_categories()


	############################9###########################################

	for item in productData:
		driver =webdriver.Firefox(options=options)
		driver.implicitly_wait(1)
		driver.get(item["url"])
		productsPageData=driver.find_elements_by_class_name("LinkProduct")
		# for x in range(5):
		for x in range(10000):
			for product in productsPageData:
				productDtailUrl=product.get_attribute("href")
				driver1 =webdriver.Firefox(options=options)
				driver1.implicitly_wait(1)
				driver1.get(productDtailUrl)
				try:
					
						try:
							prodTitle=driver1.find_element_by_class_name("ProductTitle").text
						except:
							print("")
						# print("prodTitle======",prodTitle)
						prodCode=driver1.find_element_by_class_name("ProductCode").find_element_by_class_name("kod").text
						# print("prodCode======",prodCode)
						prodPrice=driver1.find_element_by_class_name("ProductPriceContainer").text

						prod = ET.SubElement(products, 'product')
						productCode=ET.SubElement(prod,"code")
						productCode.text=prodCode

						category=ET.SubElement(prod, 'category')
						category.text=item["catName"]

						title=ET.SubElement(prod,"title")
						title.text=prodTitle

						price=ET.SubElement(prod,"price")
						price.text=prodPrice.split(' ')[1].split('cena')[1]

						currency=ET.SubElement(prod,"currency")
						currency.text="€"

						# ProdDescription=
						description=ET.SubElement(prod,"description")
						try:
							prodDes=driver1.find_element_by_xpath('//*[@id="ProductTabsContainer"]/div/div[2]/div[1]/div[1]/div/div[2]/div[2]/p')
							description.text=prodDes.text
						except:
							print("dont't exist description!")

						manuInfos=driver1.find_elements_by_class_name("ProductTabDescription")
						for manuInfo in manuInfos:
							try:
								manuInfoImg=manuInfo.find_element_by_tag_name("img").get_attribute("src")
								manuInfoImgUrl=ET.SubElement(prod,"manufacture")
								manufactureImg=ET.SubElement(manuInfoImgUrl,"img")
								manufactureImg.text=manuInfoImg

								manuInformation=manuInfo.find_element_by_tag_name("div").text
								manuInfoName=ET.SubElement(manuInfoImgUrl,"name")
								manuInfoName.text=manuInformation

							except:
								print("skipping...")

						


						OEMcode=ET.SubElement(prod,"OEMcode")


						images=ET.SubElement(prod,"images")
						prodImgs=driver1.find_elements_by_class_name("GeneralIMG")
						for prodImg in prodImgs:
							prodImgUrl=prodImg.get_attribute("href")
							img=ET.SubElement(images,"img")
							img.text=prodImgUrl



						stock=ET.SubElement(prod,"stock")
						try:
							stockInfo=ET.SubElement(stock,"info")
							stockInfo.text=driver1.find_elements_by_class_name("StockAll")[0].text
						except:
							print("don't exist StockINfo")

						try:
							stockInfo1=ET.SubElement(stock,"info")
							stockInfo1.text=driver1.find_elements_by_class_name("StockAll")[1].text
						except:
							print("don't exist StockINfo")

						warranty=ET.SubElement(prod,"Warranty")
						try:
							warranty.text=driver1.find_element_by_class_name("Product_Detail_Warranty_Value").find_element_by_tag_name("a").get_attribute("href") 
						except:
							print("don't exist warranty details")


						# additional_infos=ET.SubElement(prod,"additional_infos")
						# attBtn=driver1.find_element_by_xpath('//*[@id="tabTypeLink"]')
						# driver.execute_script("arguments[0].click();", attBtn)
						# time.sleep(1)

						##parameteries===
						parameteries=ET.SubElement(prod,"parameteries")

						try:
							prodParaContainer=driver1.find_element_by_class_name("ProductParametrContainer").find_elements_by_class_name("table-row3")
							for parameterItem in prodParaContainer:
								parameter=ET.SubElement(parameteries,"parameter")

								name=ET.SubElement(parameter,"name")
								name.text=parameterItem.find_elements_by_class_name("table-cell")[0].text

								value=ET.SubElement(parameter,"value")
								value.text = parameterItem.find_elements_by_class_name("table-cell")[1].text
						except:
							print("don't exist parameters!")
						attachments=ET.SubElement(prod,"attachments")
						try:
								
								attBtn=driver1.find_element_by_id("tabTypeLink")
								driver1.execute_script("arguments[0].click();", attBtn)
								time.sleep(1)
								prodAttfiles1=driver1.find_elements_by_class_name("ProductParametrContainer")
								for pp in prodAttfiles1:
									prodAttfiles=pp.find_elements_by_tag_name("a")
									for prodAttfile in prodAttfiles:
										try:
											attachfile=ET.SubElement(attachments,"attachment")
											attachfile.text=prodAttfile.get_attribute("href")
										except:
											print("attachfile getting error!")
						except:
								print("")

						##end parameteries===
						driver1.quit()

						mydata = ET.tostring(products,encoding='utf8').decode("utf8")
						myfile = open("prod.xml", "w",encoding='utf8')
						myfile.write(mydata)

						# break

					###next page click event
						
				except:
					print("")
			btn_status=False
			try:
							nextButtons=driver.find_element_by_class_name("product-image-list").find_elements_by_class_name("btn-default")
							
							for btn in nextButtons:
								try:
									# print("span--",btn.text)
									if(btn.text=="Zobraziť ďalšie"):
										try:
											clickbtn=driver.find_elements_by_class_name("pagination")[1].find_elements_by_tag_name("li")[7].find_element_by_tag_name("a")
											driver.execute_script("arguments[0].click();", clickbtn)
											# print("dflkdsjfskdjfkl",clickbtn.text)
											time.sleep(1)
											btn_status=True
										except:
											print("click error!")
								except:
									print("error")


							if btn_status==False:
								# print("btn_status false")
								break
			except:
				print("next button click error!")
				break

		driver.quit()

		print("Success!")




	


