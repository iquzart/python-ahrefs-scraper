from selenium import webdriver 
from time import sleep 
from webdriver_manager.chrome import ChromeDriverManager 

driver = webdriver.Chrome(ChromeDriverManager().install()) 

def login(cfg):
    '''
    login to ahrefs.com
    '''
    
    # setting up variables from cfg
    login_url = cfg['login_url']
    username = cfg['username']
    password = cfg['password']
    
    sign_in_button_xpath = '//button[@type="submit"]'
    

    driver.get(login_url) 
    print ("Opened URL:", login_url) 
    sleep(1) 
  
    _1_username_box = driver.find_element_by_name('email') 
    _1_username_box.send_keys(username) 
    print ("User Id entered", username) 
    sleep(2) 
  
    _2_password_box = driver.find_element_by_name('password') 
    _2_password_box.send_keys(password) 
    print ("Password entered") 
    sleep(2)

    _3_login_box = driver.find_element_by_xpath(sign_in_button_xpath) 
    _3_login_box.click() 
    print ("Login button clicked") 
    sleep(4) 

def logout(cfg):
    '''
    Log out from ahrefs.com
    '''

    sign_out_button_xpath = '//button[@class="btn css-10mx88t-button css-ywo6x6-control css-1c6s9tz-toggle default btn--default inverse"]'
    sign_out_xpath = '//a[@href="/user/logout"]'

    _9_expand_sign_out = driver.find_element_by_xpath(sign_out_button_xpath)
    _9_expand_sign_out.click()
    print ("Signing Out")
    sleep(1)

    _10_sign_out = driver.find_element_by_xpath(sign_out_xpath)
    _10_sign_out.click()
    print ("Sign out completed")
    sleep(5)
 
    driver.quit() 
    print("Closed Chrome")

def export_csv(target_url):
    '''
    Download CSV files based on the URLs provided in config.yaml
    '''

    target_url_box_xpath = '//input[@placeholder="Domain or URL"]'
    target_url_submit_button_xpath = '//button[@class="btn css-10mx88t-button css-184i2sh-button default btn--primary"]'
    ur_xpath = '//div[@id="UrlRatingContainer"]/span'
    dr_xpath = '//div[@id="DomainRatingContainer"]/span'
    backlinks_link_xpath = '//a[@data-nav-type="se_backlinks"]'
    home_link_xpath = '//a[@href="/dashboard"]'

    _4_url_box = driver.find_element_by_xpath(target_url_box_xpath) 
    _4_url_box.send_keys(target_url)
    print ("added url  {} on the search box".format(target_url)) 
    sleep(2)

    _5_submit_button = driver.find_element_by_xpath(target_url_submit_button_xpath) 
    _5_submit_button.click() 
    print("initiated URL search")
    sleep(3) 

    ur_value = driver.find_element_by_xpath(ur_xpath)
    ur = ur_value.text
    print("Stored UR Value: ", ur)


    dr_value = driver.find_element_by_xpath(dr_xpath)
    dr = dr_value.text
    print ("Stored DR Value: ", dr)

    _6_backlinks_link = driver.find_element_by_xpath(backlinks_link_xpath) 
    _6_backlinks_link.click() 
    print ("Backlinks link clicked") 
    sleep(3) 
  
    _7_export_button = driver.find_element_by_id("export_button")
    _7_export_button.click() 
    print ("csv_export_button clicked") 
    sleep(3)

    _8_start_export = driver.find_element_by_id('start_export_button')
    _8_start_export.click()
    print ("Export Completed")
    sleep(5)

    _9_home_link = driver.find_element_by_xpath(home_link_xpath) 
    _9_home_link.click() 
    print ("Home link clicked") 
    sleep(3) 
    
    # Return DR and UR values
    return dr, ur


