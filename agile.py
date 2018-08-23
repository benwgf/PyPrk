from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
import pandas
import re

class Agile:
    __driver=None
    def __init__(self):
        options = Options()
        options.add_argument("start-maximized")
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--no-sandbox") 
        self.__driver = webdriver.Chrome(chrome_options=options) #chrome_options=options
        #self.__driver = webdriver.PhantomJS()

        self.__driver.implicitly_wait(5)
        self.__driver.get("https://plm.sandisk.com/Agile/default/login-cms.jsp")

        self.__driver.find_element_by_id("j_username").clear()
        self.__driver.find_element_by_id("j_username").send_keys("12706")    #your windows username
        self.__driver.find_element_by_id("j_password").clear()
        self.__driver.find_element_by_id("j_password").send_keys("xxxxxxx")  #your windows passort
        self.__driver.find_element_by_id("loginspan").click()

        for handle in self.__driver.window_handles:
            self.__driver.switch_to_window(handle)

        self.__driver.maximize_window()
        time.sleep(5)

        try:
            WebDriverWait(self.__driver, 10).until(
                    EC.presence_of_element_located((By.ID, "gridwrapper_NOTIFICATION_TABLE_GRID"))
                )
        except Exception as e:
            print('TimeoutException')


    def GetAgileInfo(self,item):

        print("looking for %s"%item)
        agileinfo={}
        QUICKSEARCH_STRING=self.__driver.find_element_by_id('QUICKSEARCH_STRING')
        webdriver.ActionChains(self.__driver).move_to_element(QUICKSEARCH_STRING).click(QUICKSEARCH_STRING).perform()
        QUICKSEARCH_STRING.clear()
        time.sleep(1)
        QUICKSEARCH_STRING.send_keys(item)

        self.__driver.find_element_by_id("top_simpleSearch").click()

        try:
            WebDriverWait(self.__driver, 5).until(EC.presence_of_element_located((By.LINK_TEXT,item))).click()
        except Exception as e:
            time.sleep(1)
            pass
        try:
            WebDriverWait(self.__driver, 5).until(EC.presence_of_element_located((By.ID,'revSelectName')))
            changeList = self.__driver.find_element_by_id('revSelectName')
            all_options = changeList.find_elements_by_tag_name("option")
            for option in all_options:
                change=option.get_attribute("innerHTML")
                m = re.search(r'(\w+)(&nbsp;)+(TRC-\d+)', change)
                if m is not None:
                    agileinfo['rev']= m.group(1)
                    agileinfo['trc']= m.group(3)
                    break
            self.__driver.find_element_by_link_text("Where Used").click()
            wherelisttable= self.__driver.find_element_by_id("ITEMTABLE_WHERELIST")
            wherelist=[]

##            for row in wherelisttable.find_elements_by_css_selector("tr.GMDataRow"):
##                cells=row.find_elements_by_tag_name("td")
##                if('OBS' not in cells[7].get_attribute("innerHTML"))
##                    print(cell.get_attribute("innerHTML"))
           
            for link in wherelisttable.find_elements_by_class_name("image_link"):
                wherelist.append(link.get_attribute("innerHTML"))
            agileinfo['wherelist']=  ','.join(str(w) for w in wherelist)
            
            self.__driver.find_element_by_partial_link_text("Attachments").click()
            attachmenttable= self.__driver.find_element_by_id("ATTACHMENTS_FILELIST")
            attachments=[]           
            for link in attachmenttable.find_elements_by_class_name("image_link"):
                txt=link.get_attribute("innerHTML")
                if '.zip' in txt or '.tgz' in txt or '.bot' in txt :
                    attachments.append(txt)
                
            agileinfo['attachments']=  ','.join(str(a) for a in attachments)
            
        except Exception as e:
            print(e)

        return agileinfo;
    def GetTrcInfo(self,trc):
        
        trcinfo={}
        
        QUICKSEARCH_STRING=self.__driver.find_element_by_id('QUICKSEARCH_STRING')
        webdriver.ActionChains(self.__driver).move_to_element(QUICKSEARCH_STRING).click(QUICKSEARCH_STRING).perform()
        QUICKSEARCH_STRING.clear()
        time.sleep(1)
        QUICKSEARCH_STRING.send_keys(trc)

        self.__driver.find_element_by_id("top_simpleSearch").click()
        WebDriverWait(self.__driver, 5).until(EC.presence_of_element_located((By.ID,'dms')))
        fs= self.__driver.find_elements_by_class_name("side_by_side_text")
        title=fs[0].find_elements_by_tag_name("dd")[4].get_attribute("innerHTML").replace('<br>','\r\n').replace('&nbsp;',' ').replace('"','')
        reason=fs[0].find_elements_by_tag_name("dd")[5].get_attribute("innerHTML").replace('<br>','\r\n').replace('&nbsp;',' ').replace('"','')
        testtime=fs[2].find_elements_by_tag_name("dd")[13].get_attribute("innerHTML").replace('<br>','\r\n').replace('&nbsp;',' ').replace('"','')

        trcinfo['title']=title
        trcinfo['reason']=reason
        trcinfo['testtime']=testtime
        return trcinfo
                
    def Quit(self):
        self.__driver.quit()
               
        
if __name__ == "__main__":
 
    agile=Agile()
    df=pandas.read_excel("program readiness.xlsx")
    try:
        for index, row in df.iterrows():
            agileinfo=agile.GetAgileInfo(row["agile"].strip())
            if(len(agileinfo)!=0):
                df.loc[index,"rev"]=agileinfo['rev']
                df.loc[index,"trc"]=agileinfo['trc']
                df.loc[index,"filename"]=agileinfo['attachments']
                df.loc[index,"whereuse"]=agileinfo['wherelist']
                trcinfo=agile.GetTrcInfo(agileinfo['trc'])
                df.loc[index,"desc"]=trcinfo['title']
                df.loc[index,"reason"]=trcinfo['reason']
                df.loc[index,"tt"]=trcinfo['testtime']
            
    except Exception as e:
            print(e)
    finally:
         df.to_excel("program readiness.xlsx",index=False)
         agile.Quit()
