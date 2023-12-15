from selenium import webdriver
import time
import xlsxwriter
from tkinter import *
from selenium.webdriver.common.by import By
from scrapy import Selector
from selenium.common.exceptions import InvalidSessionIdException
city_geo_urns = {
    "Miami-Fort Lauderdale area": "90000056",
    "Greater Orlando": "90000596",
    "Jacksonville, FL": "100868799",
    "Boca Raton, FL": "103462227",
    "West Palm Beach, FL": "102574077",
    "Greater Tampa Bay Area": "90000828",
    "Fort Myers": "104948205",
    "Cape Coral Metro Area": "90000270",
    "Bradenton": "104210745",
    "North Port-Sarasota-Area": "90000751",
    "Pensacola Metropolitan Area": "90000608",
    "Naples, FL": "106919338",
    "Greater Palm Bay-Melbourne-Titusville Area": "90010479",
    "Pompano Beach, FL": "105375326",
    "Lakeland, FL": "105946785",
    "Melbourne, FL": "106033654",
    "Jupiter, FL": "100638551",
    "Crestview-Fort Walton Beach-Destin Area": "90009453",
    "Daytona Beach, FL": "104779438",
    "Greater Sebring-Avon Park Area": "90010481",
    "Port St Lucie, FL": "106921907",
    "Tallahassee Metro Area": "90000824",
}
# note only gets emails and phone numbers when subscribed to sales navigator
class Linkedin():
    def getData(self):
        service = webdriver.ChromeService(executable_path = '/Users/michaelpichardo/Downloads/chromedriver-mac-arm64/chromedriver')
        driver = webdriver.Chrome(service=service)
        driver.get('https://www.linkedin.com/login')
        
        driver.find_element(By.ID, 'username').send_keys('USER') #Enter username of linkedin account here
        time.sleep(2)
        driver.find_element(By.ID, 'password').send_keys('PASS')  #Enter Password of linkedin account here
        time.sleep(2)
        driver.find_element(By.XPATH, "//*[@type='submit']").click()
        time.sleep(40)
        #*********** Search Result ***************#
        geo_urn ="90000056"
        search_key = "software engineers" # Enter your Search key here to find people
        key = search_key.split()
        keyword = ""

        for key1 in key:
            keyword = keyword + str(key1).capitalize() +"%20"
        keyword = keyword.rstrip("%20")
            
        global data
        data = []

        for no in range(1,100):
            start = "&page={}".format(no) 
            search_url = "https://www.linkedin.com/search/results/people/?geoUrn=%5B{}%5D&keywords={}&origin=SUGGESTION{}".format(geo_urn, keyword, start)
            driver.get(search_url)
            driver.maximize_window()
            for scroll in range(2):
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)            
            search = Selector(text=(driver.page_source))
            peoples = search.css('[class="app-aware-link  scale-down "]').xpath('@href').getall()
            print(peoples) 
            count = 0
            print("Going to scrape Page {} data".format(no))
            time.sleep(2)
            # [:3] if want to limit it to 3 people for testing
            for profile_url in peoples:
                count+=1
                item = dict()
                item['profile_url'] = profile_url
                if count%1==0:
                    
                    driver.get(profile_url)
                    time.sleep(3)
                    
                    # #********** Profile Details **************#
                    page = Selector(text=(driver.page_source))
                    #from the search results page`` 
                    try:
                        item['name'] = page.css('h1.text-heading-xlarge::text').get('').strip()
                    except:
                        item['name'] = 'None'
                    try:
                        item['title'] = page.xpath('//div[@id="experience"]/following-sibling::div[@class="pvs-list__outer-container"]/ul/li//div[contains(@class,"t-bold")]/span[1]/text()').get('').strip()
                    except:
                        item['title'] = 'None'
                    try:
                        item['heading'] = page.css(
                            'div.text-body-medium.break-words::text').get('').strip()
                    except:
                        item['heading'] = 'None'
                    try:
                        item['location'] = page.css('span.inline::text').get('').strip()
                    except:
                        item['location'] = 'None'
                    
                    # import ipdb;ipdb.set_trace()
                    #*******  Contact Information **********#
                    time.sleep(2)
                    contact_info_link = page.css('#top-card-text-details-contact-info').xpath('@href').get('')
                    contact_info = "https://www.linkedin.com" + contact_info_link
                    driver.get(contact_info)
                    print(contact_info)
                    info = Selector(text=(driver.page_source))
                    time.sleep(4)
                    # Select all h3 elements within the specified CSS selector
                    h3_elements = info.css('section.pv-contact-info__contact-type > h3')
                    for h3 in h3_elements:
                        # Extract the text label of the h3 element
                        text_label = h3.xpath('normalize-space()').extract_first().lower()
                        # import ipdb;ipdb.set_trace()
                        try:
                            # Extract the following sibling elements using XPath to dynamically associate with the label
                            if text_label in ['website']:
                                # websites = h3.xpath('/following-sibling::ul/li/a/@href').getall()
                                # website = ', '.join(websites) if websites else 'None'
                                website = info.css('section > div > section:nth-child(2) > ul > li > a').xpath('@href').extract_first()
                                item['website'] = website
                            elif text_label == 'phone':
                                item['phone'] = info.css('section > div > section:nth-child(3) > ul > li').xpath('normalize-space(string())').extract_first()
                            elif text_label == 'email':
                                email = info.css('section > div > section:nth-child(3) > div > a').xpath('@href').get('').replace('mailto:', '').strip()
                                item['email'] = email
                        except InvalidSessionIdException as e:
                            print(f"InvalidSessionIdException: {e}")
                        except Exception as e:
                            print(f"An error occurred while extracting '{text_label}': {e}")
                        data.append(item)
                    print(item) 
            print("!!!!!! Data scrapped !!!!!!")
                
            driver.close()
    def writeData(self):
        workbook = xlsxwriter.Workbook("linkedin-search-data.xlsx")
        worksheet = workbook.add_worksheet('Peoples')
        bold = workbook.add_format({'bold': True})
        # headers
        worksheet.write(0,0,'profile_url',bold)
        worksheet.write(0,1,'name',bold)
        worksheet.write(0,2,'title',bold)
        worksheet.write(0,3,'heading',bold)
        worksheet.write(0,4,'location',bold)
        worksheet.write(0,5,'website',bold)
        worksheet.write(0,6,'phone',bold)
        worksheet.write(0,7,'email',bold)
        # Start writing data from the first row after the headers (row index 1)
        for i in range(1,len(data)+1): 
            try:
                worksheet.write(i, 0, data[i]['profile_url'])  # Write 'profile_url' to the worksheet
            except:
                pass
            try:
                worksheet.write(i, 1, data[i]['name'])  # Write 'profile_url' to the worksheet
            except:
                pass
            try:
                worksheet.write(i, 2, data[i]['title'])  # Write 'title' to the worksheet
            except:
                pass
            try:
                worksheet.write(i, 3, data[i]['heading'])  # Write 'heading' to the worksheet
            except:
                pass
            try:
                worksheet.write(i, 4, data[i]['location'])  # Write 'location' to the worksheet
            except:
                pass
            try:
                worksheet.write(i, 5, data[i]['website'])  # Write 'website' to the worksheet
            except:
                pass
            try:
                worksheet.write(i, 6, data[i]['phone'])  # Write 'phone' to the worksheet
            except:
                pass
            try:
                worksheet.write(i, 7, data[i]['email'])  # Write 'email' to the worksheet
            except:
                pass
        
        workbook.close()

    def start(self):
        self.getData()
        self.writeData()
if __name__ == "__main__":
    obJH = Linkedin()
    obJH.start()
