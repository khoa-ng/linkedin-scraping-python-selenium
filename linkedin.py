  
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import time
from selenium.webdriver.common.by import By
from datetime import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from scrapy.selector import Selector
import pandas as pd
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
import settings

class BrowserAutomation(object):

    def __init__(self):
       
        capabilities = DesiredCapabilities.PHANTOMJS
        capabilities["phantomjs.page.settings.userAgent"] = (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_4) AppleWebKit/600.7.12 (KHTML, like Gecko) Version/8.0.7 Safari/600.7.12"
        )   

        options = webdriver.ChromeOptions()
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--no-sandbox')
        options.add_argument("--test-type")
        options.add_argument("--start-maximized")
        executable_path=settings.path_of_chrome_driver
        self.driver = webdriver.Chrome(executable_path=executable_path,desired_capabilities=capabilities)
        self.driver.implicitly_wait(5) # seconds
        self.driver.set_window_size(1920,1080)
        self.job=[]
        self.education=[]
        self.profile_link=[]
        self.name=[]
        self.company=[]
        self.idofloadeddivs=[]
        self.backspace=0

    def launch_browser(self):

        print ('Launching Browser')
        self.driver.get("https://www.linkedin.com")
        time.sleep(5)
        print ("Browser launch successfully.")
        
    def login_to_LinkedIn(self):
        print ('Logging to LinkedIn')
        self.driver.find_element_by_xpath('//input[@class="login-email"]').send_keys(settings.linkedin_email)
        self.driver.find_element_by_xpath('//input[@class="login-password"]').send_keys(settings.linkedin_password)
        self.driver.find_element_by_xpath('//input[@class="login submit-button"]').click()        
        time.sleep(3)
    def start(self):
        
        df=pd.read_excel(settings.path_of_input_file)
        #print(len(df))
        count=0
        counti=0
        for i in range (0, len(df)):
            names=df[df.columns[1]][i]
            names=names.splitlines()
            for name in names:
                print('#'*70)                
                count+=1
                #print(count)
                if name.replace(' ','').isalpha():
                    name=name.replace('Mr ','').replace('Ms ','')
                    company=df[df.columns[0]][i]
                    company=company.replace('\n',' ')
                    print('Input Name: ' +str(name))
                    print('Input Company: '+str(company))
                    self.enter_search(name,company)
                    print('#'*70)                                     
                else:                    
                    print('Incorrect Name. Name must only include Alphabest')
                    pass

    def enter_search(self,name,company):
      
      self.company.append(company)
      self.name.append(name)
      for i in range(0, self.backspace):
       self.driver.find_element_by_xpath('//input[@role="combobox"]').send_keys(Keys.BACKSPACE)
      self.driver.find_element_by_xpath('//input[@role="combobox"]').send_keys(name)
      time.sleep(3)
      self.driver.find_element_by_xpath('//input[@role="combobox"]').send_keys(Keys.RETURN)
      time.sleep(3)
      try:
          self.driver.find_element_by_xpath('//button[@aria-controls="current-companies-facet-values"]').click()
      except Exception:
          pass
      else:
          self.driver.find_element_by_xpath('//div[@id="current-companies-facet-values"]//input[@role="combobox"]').send_keys(company)
          time.sleep(4)
          self.driver.find_element_by_xpath('//div[@id="current-companies-facet-values"]//input[@role="combobox"]').send_keys(Keys.RETURN)
          try:
           self.driver.find_element_by_xpath('//div[@id="current-companies-facet-values"]//button[@data-control-name="filter_pill_apply"]').click()
          except Exception:
              pass
          time.sleep(5)
          #self.driver.find_element_by_xpath('//li[@class="search-result search-result__occluded-item ember-view"][1]/div/div/div[2]/a').click()
      self.backspace=len(name)
      source=self.driver.page_source
      sel=Selector(text=source)
      link=str(sel.xpath('//li[@class="search-result search-result__occluded-item ember-view"][1]/div/div/div[2]/a/@href').extract()).strip("'[]")
      #print(link)
      if link:
          print('Result Found')
          link='https://www.linkedin.com'+link
          self.driver.get(link)          
          time.sleep(5)
          self.profile_link.append(self.driver.current_url)
          self.conn_roles()
          self.conn_education()          
      else:
          print('No Results Found for the Input')
          self.education.append('Publicly Unavailable')
          self.job.append('Publicly Unavailable')
          self.profile_link.append('')
          self.driver.get('https://www.linkedin.com')
          

    def conn_roles(self):
        time.sleep(3)
        #Scroll donwn to 800 and reach the experience section
        #get the count of present element
        #move through each element and perform following operation
        #Check if in the element more role button is present
        #Click as long as button exist
        #when reached the last element check for the presence of button
        #if present click ited
        #change the variables, start and end.
        #start=end
        #end=count
        #start the loop again
        #Exit after finished
        #Get source and go to eduation
        
        self.driver.execute_script("window.scrollTo(0, 800);")
        time.sleep(3)
        sel=self.get_selector()
        showmorebutton=True
        self.idofloadeddivs=sel.xpath('//div[@class="pv-entity__position-group-pager ember-view"]/@id').extract()        
        countofelementsinexperiencesection=len(self.idofloadeddivs)
        if int(countofelementsinexperiencesection)>0:
            print('Jobs Found')
            start=1
            end=countofelementsinexperiencesection+1
            #print('Count of Currently Loaded Divs is'+str(end))
            while showmorebutton:
                for i in range(start,end):
                    time.sleep(1)
                    try:
                     #Find location of the element and scroll to there
                     location=self.driver.find_element_by_xpath('//div[@id="'+str(self.idofloadeddivs[i-2])+'"]').location
                     #print('Location of ELement'+str(i)+str(location['y']))
                    except Exception:
                        pass
                    else:
                        scrollto=location['y']
                        self.driver.execute_script("window.scrollTo(0, "+str(scrollto)+");")                     
                    try:
                        #move through each block and check if show more role button present
                        self.driver.find_element_by_xpath('//div[@id="'+str(self.idofloadeddivs[i-2])+'"]//button').click()
                        
                    except Exception:
                        pass
                    else:
                        #print('More Role button Found and clicked')
                        pass
                try:
                 self.driver.find_element_by_xpath('//section[@id="experience-section"]//button[@class="pv-profile-section__see-more-inline pv-profile-section__text-truncate-toggle link"]').click()
                except Exception:
                    #print(e)
                    showmorebutton=False
                else:
                    time.sleep(3)
                    #print('Clicked the Show more button')
                    del self.idofloadeddivs[:]
                    sel=self.get_selector()                    
                    self.idofloadeddivs=sel.xpath('//div[@class="pv-entity__position-group-pager ember-view"]/@id').extract()                 
                    start=end
                    end=len(self.idofloadeddivs)+1
                    #print('New end is:'+str(end)+'New start is:'+str(start))
            sel=self.get_selector()
            data=''
            for i in range(0, len(self.idofloadeddivs)):
                                    #gotoeachblockandcheckifthedivhasmore
                                    #Ifmoreblocksusethiselseusetheotherone
                                    #moreroles=//button[contains(text(),"role")]
                numberoftitlesincurrentblock=str(sel.xpath('count(//div[@id="'+str(self.idofloadeddivs[i])+'"]//li/ul/li)').extract()).strip("'[]").replace('.0','')
                if int(numberoftitlesincurrentblock)>0:####GET THE TITLES IN A NESTED BLOCK
                    #print('Inside the Nested Block')
                    company=str(sel.xpath('//div[@id="'+str(self.idofloadeddivs[i])+'"]/li/a/div/div[2]/h3//span[2]//text()').extract()).strip("'[]")
                    #print(company)
                    for z in range(1,int(numberoftitlesincurrentblock)+1):
                        role=str(sel.xpath('//div[@id="'+str(self.idofloadeddivs[i])+'"]/li/ul/li['+str(z)+']//h3//span[2]/text()').extract()).strip("'[]")
                        #print('Title'+str(i)+role)
                        data+=company+' : '+role+' \n'
                else:
                 #print('Outside Nested')
                 role=str(sel.xpath('//div[@id="'+str(self.idofloadeddivs[i])+'"]/li/a//h3[@class="Sans-17px-black-85%-semibold"]/text()').extract()).strip("'[]")
                 #print(role)
                 company=str(sel.xpath('//div[@id="'+str(self.idofloadeddivs[i])+'"]//span[@class="pv-entity__secondary-title"]/text()').extract()).strip("'[]")
                 #print(company)
                 data+=company+' : '+role+' \n'
            self.job.append(data)
           
        else:
                print('No Jobs Found')
                self.job.append('Publicly Unavailable')
    def conn_education(self):             
        time.sleep(3)
        y=2000
        self.driver.execute_script("window.scrollTo(0, "+str(y)+");")
        for i in range(0,5):
             try:
                 y=y+300
                 self.driver.find_element_by_xpath('//section[@id="education-section"]//button[@class="pv-profile-section__see-more-inline pv-profile-section__text-truncate-toggle link"]').click()
                 time.sleep(3)
                 self.driver.execute_script("window.scrollTo(0, "+str(y)+");")                 
             except Exception:
       
                 break
        time.sleep(3)
        source=self.driver.page_source
        sel=Selector(text=source)
        count=str(sel.xpath('count(//section[@id="education-section"]/ul/li)').extract()).strip("'[]").replace('.0','')
        if int(count)>0:
            print('Education Data Present')
            data=''
            for i in range(1,int(count)+1):
                school=str(sel.xpath('//li['+str(i)+']//h3[@class="pv-entity__school-name Sans-17px-black-85%-semibold"]/text()').extract()).strip("'[]")
                #print(school)
                degree=str(sel.xpath('//li['+str(i)+']//p/span[@class="pv-entity__comma-item"]/text()').extract()).strip("'[]")
                degree=''.join(degree)
                #print(degree)
                data+=school+' : '+degree+' \n'
            self.education.append(data)
        else:
            print('No Education Data')
            self.education.append('Publicly Unavailable')
			
        out_path=settings.path_of_output_file
        list3=list(zip(self.company,self.name,self.education,self.job,self.profile_link))
        df=pd.DataFrame(data=list3,columns=['Company','Name','Education','Job','Profile Link'])                   
        writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
        df.to_excel(writer,index=False,header=True)#always overwrites, ile can be saved in any format, docx, json, xml,etc
        writer.save()
        writer.close()                                   
       

    def save_info(self):
        out_path=settings.path_of_output_file
        list3=list(zip(self.company,self.name,self.education,self.job,self.profile_link))
        df=pd.DataFrame(data=list3,columns=['Company','Name','Education','Job','Profile Link'])                   
        writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
        df.to_excel(writer,index=False,header=True)#always overwrites, ile can be saved in any format, docx, json, xml,etc
        writer.save()
        writer.close()
    def get_selector(self):
            source=self.driver.page_source
            sel=Selector(text=source)
            return sel
        
    def run_automation(self):
        print ("=============================")
        print ("     Starting automation     ")
        print ("=============================")
        self.launch_browser()
        self.login_to_LinkedIn()
        self.start()
        self.save_info()
        self.driver.quit()

print ('='*100)
print ("Script Started At : "  + datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
print ('='*100)
try:
    browser = BrowserAutomation()
    browser.run_automation()
     # quit the node proc
except Exception as e:
    print ("Error occured ")
    raise
print ('='*100)
print ("Script End At : " +  datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
print ('='*100)





