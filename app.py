#from webui import WebUI
from flask import Flask, render_template, request, make_response
from config import Config
import requests
import csv
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
#import pdfkit

app = Flask(__name__)
#ui = WebUI(app, debug=True) # Create a WebUI instance

@app.route('/')
@app.route('/scrape', methods=['GET', 'POST'])
def scrape():
    if request.method == 'GET':
        return render_template('index.html')

    if request.method == 'POST':
        record = request.form['record']
        browser = webdriver.Chrome()
        # Login url
        url = 'https://www.mexicanautoinsurance.com/user/login'
        # Get the url
        browser.get(url)
        # Grab username
        browser.find_element_by_name('identity').send_keys(Config.LEWIS_USERNAME)
        # Grab password
        browser.find_element_by_name('credential').send_keys(Config.LEWIS_PASSWORD)
        # Submit and login
        browser.find_element_by_name('submit').click()

        browser.implicitly_wait(.5)
        wait = WebDriverWait(browser, .5)
        # Basic fields to start and test with
        url = 'https://www.mexicanautoinsurance.com/staff/homeowners-insurance#/details/'
        client = [[]]
        data = []
        # Grab first two home quotations by id
        # ===========================================for page in range(record):
        # empty array to store td's for each client
        browser.get(url + str(record))
        # Wait until table is visible on each page of home details to scrape
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'table')))
        # Create the soup
        soup = BeautifulSoup(browser.page_source, 'lxml')
        # Grab each table
        
        element = soup.select('table td + td')
        element = [el.text for el in element]
        client.append(element)

        browser.get('https://www.mexicanautoinsurance.com/staff/homeowners-insurance#/')

    with open('home_data.csv', 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter = ',')
        # Select the fields and write one row with headers
        field = soup.select('table tbody tr td')[::2]
        fields = [el.text for el in field]
        fields = fields[0:36]
        fields.append('Outdoor Property')
        fields.append('Other Structures')
        csvwriter.writerow(fields)

        # Find Outdoor Property
        op = 0
        outdoor_property = soup.select('table')[4].select('tbody tr td + td')
        if outdoor_property is None:
            outdoor_property = [int(el.text.replace('$', '').replace(',', '')) for el in outdoor_property]
            op = sum(outdoor_property)

        # Find Other Structures
        structures = soup.select('table')[5].select('tbody tr td + td')
        ot = 0
        if structures is None:
            structures = [int(el.text.replace('$', '').replace(',', '')) for el in structures]
            ot = sum(structures)
        # Total Other Structures
        for c in client:
            c = c[0:36]
            c.append(op)
            c.append(ot)
            csvwriter.writerows([c])
        
    with open('home_data.csv', 'r') as data:
        csv_data = csv.reader(data)
        r = [el for el in csv_data]
        first_name = r[2][0]
        last_name = r[2][2]
        email = r[2][3]
        phone_1 = r[2][4][5:8]
        phone_2 = r[2][4][10:13]
        phone_3 = r[2][4][14:18]
        street = r[2][5] 
        city = r[2][6]
        state = r[2][7]
        zip_code = r[2][8][0:5]
        mex_street = r[2][9]
        mex_city = r[2][10]
        mex_state = r[2][11]
        mex_zip_code = r[2][12][0:5]
        type_of_risk = r[2][15]
        type_of_property = r[2][16]
        use_of_property = r[2][17]
        underground_floors = r[2][18]
        roof_type = r[2][19]
        main_dwelling = r[2][20]
        storm_shutters = r[2][22]
        adjacent_structures = r[2][23]
        number_of_stories = r[2][24][0]
        coast_line = r[2][25]
        sea_level = r[2][26]
        dwelling_amount = r[2][27]
        dwelling_amount = dwelling_amount.replace('$', '').replace(',', '')
        burglary = r[2][28]
        burglary = burglary.replace('$', '').replace(',', '')
        contents = r[2][29]
        contents = contents.replace('$', '').replace(',', '')
        loss_of_rents = r[2][30]
        liability = r[2][31]
        liability = liability.replace('$', '').replace(',', '')
        valuable_items = r[2][32]
        valuable_items = valuable_items.replace('$', '').replace(',', '')
        money = r[2][33]
        money = money.replace('$', '').replace(',', '')
        breakage = r[2][34]
        breakage = breakage.replace('$', '').replace(',', '')
        electric = r[2][35]
        electric = electric.replace('$', '').replace(',', '')
        other_structures = r[2][37]
        other_structures = int(other_structures)
        outdoor_property = r[2][36]
        outdoor_property = int(outdoor_property)

        client = {
            'fn': first_name,
            'ln': last_name,
            'email': email,
            'phone_1': phone_1,
            'phone_2': phone_2,
            'phone_3': phone_3,
            'street': street,
            'city': city,
            'state': state,
            'zip_code': zip_code,
            'mex_street': mex_street,
            'mex_city': mex_city,
            'mex_state': mex_state,
            'mex_zip_code': mex_zip_code,
            'type_of_risk': type_of_risk,
            'type_of_property': type_of_property,
            'use_of_property': use_of_property,
            'underground_floors': underground_floors,
            'roof_type': roof_type,
            'main_dwelling': main_dwelling,
            'storm_shutters': storm_shutters,
            'adjacent_structures': adjacent_structures,
            'number_of_stories': number_of_stories,
            'coast_line': coast_line,
            'sea_level': sea_level,
            'dwelling_amount': dwelling_amount,
            'contents': contents,
            'loss_of_rents': loss_of_rents,
            'liability': liability,
            'valuable_items': valuable_items,
            'burglary': burglary,
            'money': money,
            'breakage': breakage,
            'electric': electric,
            'other_structures': other_structures,
            'outdoor_property': outdoor_property
        }
        return render_template('results.html', client=client)
    browser.close()

@app.route('/update-quote', methods=['GET', 'POST'])
def update():
    if request.method == 'POST':
        browser = webdriver.Chrome()

        url = 'http://protec.abcdelseguro.com/Account/Login?localAppReturn=%2fLogin.aspx&app=10&test=http%3a%2f%2fhomeowner.abcdelseguro.com'
        browser.get(url)
        browser.find_element_by_name('username').send_keys(Config.MAPFRE_USERNAME)
        browser.find_element_by_name('password').send_keys(Config.MAPFRE_PASSWORD)
        browser.find_element_by_xpath("//input[@value='Login']").click()        

        quote_url = 'http://homeowner.abcdelseguro.com/Application/Load.aspx'
        browser.get(quote_url)

        quote_num = request.form['quote_num']
        browser.find_element_by_name('ctl00$workloadContent$txtLoad').send_keys(quote_num)
        browser.find_element_by_xpath("//input[@value='Load']").click()

        r = request.form
        # Form fields
        first_name = r['first_name']
        last_name = r['last_name']
        email = r['email']
        # phone is not required if email is present
        street = r['street']
        city = r['city']
        state = r['state']
        zip_code = r['zip_code']
        mex_street = r['mex_street']
        mex_city = r['mex_city']
        mex_state = r['mex_state']
        mex_zip_code = r['mex_zip_code']
        type_of_risk = r['type_of_risk'].upper()
        type_of_property = r['type_of_property'].upper()
        use_of_property = r['use_of_property'].upper()
        underground_floors = r['underground_floors']
        roof_type = r['roof_type'].upper()
        main_dwelling = r['main_dwelling'].upper()
        storm_shutters = r['storm_shutters']
        adjacent_structures = r['adjacent_structures']
        number_of_stories = r['number_of_stories']
        coast_line = r['coast_line']
        sea_level = r['sea_level']
        dwelling_amount = r['dwelling_amount']
        contents = r['contents']
        loss_of_rents = r['loss_of_rents']
        liability = r['liability']
        valuable_items = r['valuable_items']
        money = r['money']
        burglary = r['burglary']
        breakage = r['breakage']
        electric = r['electric']
        other_structures = r['other_structures']
        outdoor_property = r['outdoor_property']

        # =================================================================
        # Page 1 of Application
        # =================================================================
        
        # First Name
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtName').clear()
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtName').send_keys(first_name)
        #Last Name
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtLastName').clear()
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtLastName').send_keys(last_name)
        # Email
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtEmail').clear()
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtEmail').send_keys(email)
        # Phone - Area Code - xxx - xxxx
        
        #browser.find_element_by_name('ctl00$workloadContent$currentStep$txtAreaPhone').send_keys(phone_1)
        #browser.find_element_by_name('ctl00$workloadContent$currentStep$txtPhone1').send_keys(phone_2)
        #browser.find_element_by_name('ctl00$workloadContent$currentStep$txtPhone2').send_keys(phone_3)

        # Client Type
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlClient']/option[text()='{type_of_risk}']").click()
        # Dwelling Amount
        if dwelling_amount:
            browser.find_element_by_name('ctl00$workloadContent$currentStep$txtValue').clear()
            browser.find_element_by_name('ctl00$workloadContent$currentStep$txtValue').send_keys(dwelling_amount)
        # Contents Amount
        if not contents:
            contents = 25000
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtContents').clear()
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtContents').send_keys(contents)
        # Property Type
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlUse']/option[text()='{use_of_property}']").click()
        # Mexican Zip Code
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtZipCode').clear()
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtZipCode').send_keys(mex_zip_code)
        # Go to Next Page
        # 1. Click once to load zip code information
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()

        # =================================================================
        # Page 2 of Application
        # =================================================================

        # Rating is not dependent on page 2 information
        # Simply for underwriting! Not to get a premium figure
        '''
        if type_of_property == 'DWELLING':
            type_of_property = 'HOUSE'
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlType']/option[text()='{type_of_property}']").click()
        # roof type
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlRoof']/option[text()='{roof_type}']").click()
        # wall type
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlWall']/option[text()='{main_dwelling}']").click()
        # Storm Shutters YES
        '''
        if storm_shutters == 'Yes':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblShutters_0").click()
        # Storm shutters NO
        if storm_shutters == 'No':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblShutters_1").click()
        # Adjacent structures YES
        if adjacent_structures == 'Yes':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblAdjoining_0").click()
        # Adjacent structures NO
        if adjacent_structures == 'No':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblAdjoining_1").click()
        # Number of Stories (Clear default amount)
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtStories').clear()
        # Add desired stories from client
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtStories').clear()
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtStories').send_keys(number_of_stories)
        # Distance from water
        # Less than
        if coast_line == 'Less than 500 Meters':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblDistanceSea_0").click()
        

        # More than
        if coast_line == 'More than 500 Meters':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblDistanceSea_1").click()
        
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()

        # =================================================================
        # Page 3 of Application
        # =================================================================

        # Click HMP
        #browser.find_element_by_id("ctl00_workloadContent_currentStep_chkHMP").click()
        
        # Other Structures
        if other_structures:
            #browser.find_element_by_id('ctl00_workloadContent_currentStep_chkOtherStructures').click()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructures').clear()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructures').send_keys(other_structures)
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructures').send_keys(Keys.TAB)

        # Outdoor Property
        if outdoor_property:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_chkOutdoor').click()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoor').clear()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoor').send_keys(outdoor_property)
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoor').send_keys(Keys.TAB)
        
        # Clear liability field
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtCSL').clear()
        # Add desired liability from client
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtCSL').send_keys(liability)
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtCSL').send_keys(Keys.TAB)
        
        # Burglary
        if burglary:
            burglary = int(burglary)
            if burglary > 0:
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkBurglary').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtBurglary').clear()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtBurglary').send_keys(burglary)
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtBurglary').send_keys(Keys.TAB)

        # Money and Securities
        if money:
            money = int(money)
            if money > 2000:
                money = 2000
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkMoney').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtMoney').clear()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtMoney').send_keys(money)
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtMoney').send_keys(Keys.TAB)

        # Accidental Glass Breakage
        if breakage:
            breakage = int(breakage)
            if breakage > 0:
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkGlass').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtGlass').clear()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtGlass').send_keys(breakage)
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtGlass').send_keys(Keys.TAB)

        # Electric
        if electric:
            electric = int(electric)
            if electric > 0:
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkElectronic').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtElectronic').clear()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtElectronic').send_keys(electric) 
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtElectronic').send_keys(Keys.TAB)

        # Uncheck Family Assistance

        if browser.find_element_by_id('ctl00_workloadContent_currentStep_chkFamily').is_selected() == True:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_chkFamily').click()
        
        if outdoor_property:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoorDescription').clear()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoorDescription').send_keys('Outdoor Property Descriptions Coming Soon')
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoorDescription').send_keys(Keys.TAB)
        
        if other_structures:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructuresDescription').clear()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructuresDescription').send_keys('Other Structure Descriptions Coming Soon')
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructuresDescription').send_keys(Keys.TAB)

        #browser.implicitly_wait(.75)
        #wait = WebDriverWait(browser, .75)
        #wait.until(EC.text_to_be_present_in_element_value((By.ID, 'ctl00_workloadContent_currentStep_txtCSL'), {liability}))

        # Go on to address information and generate premium
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()

        # Policy Fee
        fee = browser.find_element_by_id('ctl00_workloadContent_currentStep_lblFee').text
        # Grab Tax
        tax = browser.find_element_by_id('ctl00_workloadContent_currentStep_lblTax').text
        # Grab Net Premium
        net = browser.find_element_by_id('ctl00_workloadContent_currentStep_lblNetPremium').text
        net = net.replace('$', '').replace(',', '')
        # Grab Total Premium
        tp = browser.find_element_by_id('ctl00_workloadContent_currentStep_lblTotal').text

        broker_fee = 0
        if (float(net) * 0.20 < 125):
            broker_fee = 125 - (float(net) * 0.15)
        elif (float(net) * 0.20 > 125):
            broker_fee = float(net) * 0.05

        broker_fee = round(broker_fee, 2)
        tcoi = tp.replace('$', '').replace(',', '')
        tcoi = float(tcoi) + broker_fee
        tcoi = round(tcoi, 2)


        text = f'Please find homeowners quotation attached for good measure. Your total premium comes to {tcoi}. Should you have any questions or need to adjust limits, please let me know.'
        def Emailer(text, subject, recipient):   
            import win32com.client as win32
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = recipient
            mail.Subject = subject
            mail.HtmlBody = text
            mail.Display(True)

        return render_template('success.html', r=r, fee=fee, tax=tax, net=net, tp=tp, tcoi=tcoi, broker_fee = broker_fee)
        #pdf = pdfkit.from_url('http://localhost:5000/get-quote', 'quotation.pdf')
        #Emailer(text, f'Mexican Homeowners Quotation for {first_name} {last_name}', f'{email}')



@app.route('/get-quote', methods=['GET', 'POST'])
def quote():
    if request.method == 'POST':
        browser = webdriver.Chrome()

        url = 'http://protec.abcdelseguro.com/Account/Login?localAppReturn=%2fLogin.aspx&app=10&test=http%3a%2f%2fhomeowner.abcdelseguro.com'
        browser.get(url)
        browser.find_element_by_name('username').send_keys('upall')
        browser.find_element_by_name('password').send_keys('mae790')
        browser.find_element_by_xpath("//input[@value='Login']").click()

        quote_url = 'http://homeowner.abcdelseguro.com/Application/GenerateQuote.aspx'
        browser.get(quote_url)
        
        r = request.form
        # Form fields
        first_name = r['first_name']
        last_name = r['last_name']
        email = r['email']
        # phone is not required if email is present
        street = r['street']
        city = r['city']
        state = r['state']
        zip_code = r['zip_code']
        mex_street = r['mex_street']
        mex_city = r['mex_city']
        mex_state = r['mex_state']
        mex_zip_code = r['mex_zip_code']
        type_of_risk = r['type_of_risk'].upper()
        type_of_property = r['type_of_property'].upper()
        use_of_property = r['use_of_property'].upper()
        underground_floors = r['underground_floors']
        roof_type = r['roof_type'].upper()
        main_dwelling = r['main_dwelling'].upper()
        storm_shutters = r['storm_shutters']
        adjacent_structures = r['adjacent_structures']
        number_of_stories = r['number_of_stories']
        coast_line = r['coast_line']
        sea_level = r['sea_level']
        dwelling_amount = r['dwelling_amount']
        contents = r['contents']
        loss_of_rents = r['loss_of_rents']
        liability = r['liability']
        valuable_items = r['valuable_items']
        money = r['money']
        burglary = r['burglary']
        breakage = r['breakage']
        electric = r['electric']
        other_structures = r['other_structures']
        outdoor_property = r['outdoor_property']

        # =================================================================
        # Page 1 of Application
        # =================================================================
        
        # First Name
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtName').send_keys(first_name)
        #Last Name
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtLastName').send_keys(last_name)
        # Email
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtEmail').send_keys(email)
        # Phone - Area Code - xxx - xxxx
        
        #browser.find_element_by_name('ctl00$workloadContent$currentStep$txtAreaPhone').send_keys(phone_1)
        #browser.find_element_by_name('ctl00$workloadContent$currentStep$txtPhone1').send_keys(phone_2)
        #browser.find_element_by_name('ctl00$workloadContent$currentStep$txtPhone2').send_keys(phone_3)

        # Client Type
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlClient']/option[text()='{type_of_risk}']").click()
        # Dwelling Amount
        if dwelling_amount:
            browser.find_element_by_name('ctl00$workloadContent$currentStep$txtValue').send_keys(dwelling_amount)
        # Contents Amount
        if not contents:
            contents = 25000
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtContents').send_keys(contents)
        # Property Type
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlUse']/option[text()='{use_of_property}']").click()
        # Mexican Zip Code
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtZipCode').send_keys(mex_zip_code)
        # Go to Next Page
        # 1. Click once to load zip code information
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()
        # 2. Click twice to continue to next page
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()

        # =================================================================
        # Page 2 of Application
        # =================================================================

        # Rating is not dependent on page 2 information
        # Simply for underwriting! Not to get a premium figure
        if type_of_property == 'DWELLING':
            type_of_property = 'HOUSE'
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlType']/option[text()='{type_of_property}']").click()
        # roof type
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlRoof']/option[text()='{roof_type}']").click()
        # wall type
        browser.find_element_by_xpath(f"//select[@name='ctl00$workloadContent$currentStep$ddlWall']/option[text()='{main_dwelling}']").click()
        # Storm Shutters YES
        if storm_shutters == 'Yes':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblShutters_0").click()
        # Storm shutters NO
        if storm_shutters == 'No':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblShutters_1").click()
        # Adjacent structures YES
        if adjacent_structures == 'Yes':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblAdjoining_0").click()
        # Adjacent structures NO
        if adjacent_structures == 'No':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblAdjoining_1").click()
        # Number of Stories (Clear default amount)
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtStories').clear()
        # Add desired stories from client
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtStories').send_keys(number_of_stories)
        # Distance from water
        # Less than
        if coast_line == 'Less than 500 Meters':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblDistanceSea_0").click()
        

        # More than
        if coast_line == 'More than 500 Meters':
            browser.find_element_by_id("ctl00_workloadContent_currentStep_rblDistanceSea_1").click()
        
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()

        # =================================================================
        # Page 3 of Application
        # =================================================================

        # Click HMP
        browser.find_element_by_id("ctl00_workloadContent_currentStep_chkHMP").click()
        
        # Other Structures
        if other_structures:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_chkOtherStructures').click()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructures').send_keys(other_structures)
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructures').send_keys(Keys.TAB)

        # Outdoor Property
        if outdoor_property:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_chkOutdoor').click()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoor').send_keys(outdoor_property)
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoor').send_keys(Keys.TAB)
        
        # Clear liability field
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtCSL').clear()
        # Add desired liability from client
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtCSL').send_keys(liability)
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtCSL').send_keys(Keys.TAB)
        
        # Burglary
        if burglary:
            burglary = int(burglary)
            if burglary > 0:
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkBurglary').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtBurglary').send_keys(burglary)
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtBurglary').send_keys(Keys.TAB)

        # Money and Securities
        if money:
            money = int(money)
            if money >= 2000:
                money = 2000
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkMoney').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtMoney').send_keys(money)
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtMoney').send_keys(Keys.TAB)

        # Accidental Glass Breakage
        if breakage:
            breakage = int(breakage)
            if breakage > 0:
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkGlass').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtGlass').send_keys(breakage)
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtGlass').send_keys(Keys.TAB)

        # Electric
        if electric:
            electric = int(electric)
            if electric > 0:
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkElectronic').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtElectronic').send_keys(electric) 
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtElectronic').send_keys(Keys.TAB)

        # Uncheck Family Assistance
        if browser.find_element_by_id('ctl00_workloadContent_currentStep_chkFamily').is_selected() == True:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_chkFamily').click()

        if outdoor_property:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoorDescription').send_keys('Outdoor Property Descriptions Coming Soon')
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoorDescription').send_keys(Keys.TAB)
        
        if other_structures:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructuresDescription').send_keys('Other Structure Descriptions Coming Soon')
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructuresDescription').send_keys(Keys.TAB)

        #browser.implicitly_wait(.75)
        #wait = WebDriverWait(browser, .75)
        #wait.until(EC.text_to_be_present_in_element_value((By.ID, 'ctl00_workloadContent_currentStep_txtCSL'), {liability}))

        # Go on to address information and generate premium
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()

        # Policy Fee
        fee = browser.find_element_by_id('ctl00_workloadContent_currentStep_lblFee').text
        # Grab Tax
        tax = browser.find_element_by_id('ctl00_workloadContent_currentStep_lblTax').text
        # Grab Net Premium
        net = browser.find_element_by_id('ctl00_workloadContent_currentStep_lblNetPremium').text
        net = net.replace('$', '').replace(',', '')
        # Grab Total Premium
        tp = browser.find_element_by_id('ctl00_workloadContent_currentStep_lblTotal').text

        broker_fee = 0
        if (float(net) * 0.20 < 125):
            broker_fee = 125 - (float(net) * 0.15)
        elif (float(net) * 0.20 > 125):
            broker_fee = float(net) * 0.05

        broker_fee = round(broker_fee, 2)
        tcoi = tp.replace('$', '').replace(',', '')
        tcoi = float(tcoi) + broker_fee
        tcoi = round(tcoi, 2)


        text = f'Please find homeowners quotation attached for good measure. Your total premium comes to {tcoi}. Should you have any questions or need to adjust limits, please let me know.'
        def Emailer(text, subject, recipient):   
            import win32com.client as win32
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = recipient
            mail.Subject = subject
            mail.HtmlBody = text
            mail.Display(True)

        
        return render_template('success.html', r=r, fee=fee, tax=tax, net=net, tp=tp, tcoi=tcoi, broker_fee = broker_fee)
        #pdf = pdfkit.from_url('http://localhost:5000/get-quote', 'quotation.pdf')
        #Emailer(text, f'Mexican Homeowners Quotation for {first_name} {last_name}', f'{email}')

'''

@app.route('/load_csv', methods=['GET', 'POST'])
def quote():
    url = 'http://protec.abcdelseguro.com/Account/Login?localAppReturn=%2fLogin.aspx&app=10&test=http%3a%2f%2fhomeowner.abcdelseguro.com'
    browser.get(url)
    browser.find_element_by_name('username').send_keys('upall')
    browser.find_element_by_name('password').send_keys('mae790')
    browser.find_element_by_xpath("//input[@value='Login']").click()

    quote_url = 'http://homeowner.abcdelseguro.com/Application/GenerateQuote.aspx'
    browser.get(quote_url)

    with open('home_data.csv', 'r') as data:
        csv_data = csv.reader(data)
        r = [el for el in csv_data]
        first_name = r[2][0]
        last_name = r[2][2]
        email = r[2][3]
        phone_1 = r[2][4][5:8]
        phone_2 = r[2][4][10:13]
        phone_3 = r[2][4][14:18]
        dwelling_amount = r[2][27]
        dwelling_amount = dwelling_amount.replace('$', '').replace(',', '')
        contents = r[2][29]
        contents = contents.replace('$', '').replace(',', '')
        mexican_zip = r[2][12][0:5]

        # =================================================================
        # Page 1 of Application
        # =================================================================
        
        # First Name
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtName').send_keys(first_name)
        #Last Name
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtLastName').send_keys(last_name)
        # Email
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtEmail').send_keys(email)
        # Phone - Area Code - xxx - xxxx
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtAreaPhone').send_keys(phone_1)
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtPhone1').send_keys(phone_2)
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtPhone2').send_keys(phone_3)
        # Client Type
        browser.find_element_by_xpath("//select[@name='ctl00$workloadContent$currentStep$ddlClient']/option[text()='OWNER']").click()
        # Dwelling Amount
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtValue').send_keys(dwelling_amount)
        # Contents Amount
        if not contents:
            contents = 25000
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtContents').send_keys(contents)
        # Property Type
        browser.find_element_by_xpath("//select[@name='ctl00$workloadContent$currentStep$ddlUse']/option[text()='PRIMARY RESIDENCE']").click()
        # Mexican Zip Code
        browser.find_element_by_name('ctl00$workloadContent$currentStep$txtZipCode').send_keys(mexican_zip)
        # Go to Next Page
        # 1. Click once to load zip code information
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()
        # 2. Click twice to continue to next page
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()

        # =================================================================
        # Page 2 of Application
        # =================================================================

        # Rating is not dependent on page 2 information
        # Simply for underwriting! Not to get a premium figure
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()

        # =================================================================
        # Page 3 of Application
        # =================================================================

        # Civil Liability
        liability = r[2][31]
        liability = liability.replace('$', '').replace(',', '')
        other_structures = r[2][37]
        other_structures = int(other_structures)
        outdoor_property = r[2][36]
        outdoor_property = int(outdoor_property)

        # Click HMP
        browser.find_element_by_id("ctl00_workloadContent_currentStep_chkHMP").click()

        # Other Structures
        if other_structures:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_chkOtherStructures').click()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructures').send_keys(other_structures)
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructures').send_keys(Keys.TAB)

        # Outdoor Property
        if outdoor_property:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_chkOutdoor').click()
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoor').send_keys(outdoor_property)
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoor').send_keys(Keys.TAB)
        
        # Clear liability field
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtCSL').clear()
        # Add desired liability from client
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtCSL').send_keys(liability)
        browser.find_element_by_id('ctl00_workloadContent_currentStep_txtCSL').send_keys(Keys.TAB)
        
        burglary = r[2][28]
        # Burglary
        if burglary:
            burglary = burglary.replace('$', '').replace(',', '')
            burglary = int(burglary)
            if burglary > 0:
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkBurglary').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtBurglary').send_keys(burglary)
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtBurglary').send_keys(Keys.TAB)

        # Money and Securities
        money = r[2][33]
        if money:
            money = money.replace('$', '').replace(',', '')
            money = int(money)
            if money > 2000:
                money = 2000
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkMoney').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtMoney').send_keys(money)
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtMoney').send_keys(Keys.TAB)

        # Accidental Glass Breakage
        breakage = r[2][34]
        if breakage:
            breakage = breakage.replace('$', '').replace(',', '')
            breakage = int(breakage)
            if breakage > 0:
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkGlass').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtGlass').send_keys(breakage)
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtGlass').send_keys(Keys.TAB)

        # Electric
        electric = r[2][35]
        if electric:
            electric = electric.replace('$', '').replace(',', '')
            electric = int(electric)
            if electric > 0:
                browser.find_element_by_id('ctl00_workloadContent_currentStep_chkElectronic').click()
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtElectronic').send_keys(electric) 
                browser.find_element_by_id('ctl00_workloadContent_currentStep_txtElectronic').send_keys(Keys.TAB)

        # Uncheck Family Assistance
        browser.find_element_by_id('ctl00_workloadContent_currentStep_chkFamily').click()
        if outdoor_property:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoorDescription').send_keys('Outdoor Property Descriptions Coming Soon')
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOutdoorDescription').send_keys(Keys.TAB)
        
        if other_structures:
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructuresDescription').send_keys('Other Structure Descriptions Coming Soon')
            browser.find_element_by_id('ctl00_workloadContent_currentStep_txtOtherStructuresDescription').send_keys(Keys.TAB)

        # Go on to address information and generate premium
        browser.find_element_by_id("ctl00_workloadContent_btnNext").click()

        # Grab Total Premium
        total_premium = browser.find_element_by_id('ctl00_workloadContent_currentStep_lblTotal')

        df = pd.read_csv('home_data.csv')
        df['Premium'] = total_premium.text.replace('$', '').replace(',', '')
        df = df[['First Name', 'Last Name', 'Phone', 'Email', 'Dwelling Amount', 'Liability', 'Premium']]
        df.to_csv('quote.csv', index=False)
'''

if __name__ == '__main__':
    app.run(debug=True)
    #ui.run()