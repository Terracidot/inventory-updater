from playwright.sync_api import sync_playwright
import time
import csv
from datetime import datetime

stock_code_list = []
stock_qty_list = []

def timer(i):
    current_time = datetime.now().strftime("%H:%M:%S")
    print(i + ' time -', current_time)

def update_stock(code, page_id):
    try:
        page1.goto(f'https://path_of_the_webpage/admin/edit-product?id={page_id}')
        page1.get_by_role('tab', name="Pricing & variations").click()
        page1.get_by_role('textbox', name="Available Stock").fill(str(stock_dict.get(code)))
        page1.get_by_role('button', name="Save").click()
        print(code, 'new quantity =', stock_dict.get(code))
    except:
        print('Error on', code)

timer('Start')

# Main function
with sync_playwright() as p:
    browser0 = p.chromium.launch(headless=True, executable_path='C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    page0 = browser0.new_page()

    print('Logging into your_website...')
    page0.goto('https://path_of_the_webpage.com/')
    page0.fill('input#ctl00_cphFullWidthDocumentContainer_LoginControl_txtUserName_vertical', 'your_username')
    page0.fill('input#ctl00_cphFullWidthDocumentContainer_LoginControl_txtPassword_vertical', 'your_password')
    page0.get_by_role('link', name="Login").click()
    print('Sage login successful!')

    page0.goto('https://path_of_the_webpage.com/')
    page0.get_by_role('link', name='View Report').click()

    table_data = page0.locator('#reportContent')
    table_data.wait_for()

    # Get rows in HTML table
    rows = table_data.get_by_role('row')
    print('Number of stock items:', rows.count())
    print('Appending stock items...')

    # iterate through all rows and append to list
    for i in rows.all():
        try:
            stock_code_list.append(i.locator('//td[1]').inner_text())
            print("Appended:", i.locator('//td[1]').inner_text())
        except:
            print('Could not append', i.locator('//td[1]').inner_text(), 'Setting to NO CODE instead')
            stock_code_list.append('NO CODE')
        try:
            stock_qty_list.append(int(float(i.locator('//td[10]').inner_text())))
            # print(int(float(i.locator('//td[10]').inner_text())))
        except:
            print('No quantity found for', (i.locator('//td[1]').inner_text()), 'Setting to 0 instead')
            stock_qty_list.append(0)
            
    print('Append of stock items completed!')

    stock_dict = dict(zip(stock_code_list, stock_qty_list))

    # Save to csv file
    print('Writing stock info to csv...')
    
    with open('stock_list.csv', 'w+') as csv_file:  
        writer = csv.writer(csv_file)
        for dict_key, dict_value in stock_dict.items():
            writer.writerow([dict_key, dict_value])   

    print('Writing to csv completed.')

    browser0.close()

    browser1 = p.chromium.launch(headless=True, executable_path='C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    page1 = browser1.new_page()

    print('Logging into Website...')

    page1.goto('https://path_of_the_webpage.com/user/sign-in')
    page1.get_by_label("Email").fill("your_username")
    page1.get_by_label("Password").fill("your_password")
    page1.get_by_label("Log In").click()

    print('Login successful!')

    page1.get_by_role('main').get_by_role('link', name="Shop").click()
    page1.get_by_role('link', name="Products").click()

    print('Starting stock quantity update...')

    update_stock('stock_code_0', 8687)
    update_stock('stock_code_1', 10228)
    update_stock('stock_code_2', 10254)
    update_stock('stock_code_3', 10255)
    update_stock('stock_code_4', 10229)
    update_stock('stock_code_5', 10668)
    update_stock('stock_code_6', 10669)
    update_stock('stock_code_7', 10252)
    update_stock('stock_code_8', 10209)
    update_stock('stock_code_9', 10222)
    
    print('Stock update complete! Closing in 2 minutes...')

    browser1.close()

    timer('End')

    time.sleep(120)
