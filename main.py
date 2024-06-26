from playwright.sync_api import sync_playwright
import time
import csv
from datetime import datetime

stock_code_list = []
stock_qty_list = []
count = 0

def timer(i):
    current_time = datetime.now().strftime("%H:%M:%S")
    print(i + ' time -', current_time)

def update_stock(code, page_id):
    try:
        page1.goto(f'https://soundcom-shop.spaza.co/admin/edit-product?id={page_id}')
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

    print('Logging into Sage...')
    page0.goto('https://accounting.sageone.co.za/Landing/ResellerDefault.aspx')
    page0.fill('input#ctl00_cphFullWidthDocumentContainer_LoginControl_txtUserName_vertical', 'matthew@soundcom.co.za')
    page0.fill('input#ctl00_cphFullWidthDocumentContainer_LoginControl_txtPassword_vertical', 'P@$$w0rd2314')
    page0.get_by_role('link', name="Login").click()
    print('Sage login successful!')

    page0.goto('https://accounting.sageone.co.za/Reports/HTMLReportFilter.aspx?encqry=POVm97O0+iW4ekO2VNbRHB7RwBkSyDgQ+EU9o3OiQjdJRwHFgKPoVh8Rs5P82XZ4')
    page0.get_by_role('link', name='View Report').click()

    table_data = page0.locator('#reportContent')
    table_data.wait_for()

    # Get rows in HTML table
    rows = table_data.get_by_role('row')
    print('Number of stock items:', rows.count())
    print('Appending stock items...')
    
    for i in rows.all():
        # count += 1
        # print('row',count)
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
            
        # if count == 445: break

    print('Append of stock items completed!')

    stock_dict = dict(zip(stock_code_list, stock_qty_list))

    # Save dictionary to csv file
    print('Writing stock info to csv...')
    
    with open('stock_list.csv', 'w+') as csv_file:  
        writer = csv.writer(csv_file)
        for dict_key, dict_value in stock_dict.items():
            writer.writerow([dict_key, dict_value])   

    print('Writing to csv completed!')

    browser0.close()

    browser1 = p.chromium.launch(headless=True, executable_path='C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
    page1 = browser1.new_page()

    print('Logging into Spaza...')

    page1.goto('https://www.spaza.co/user/sign-in')
    page1.get_by_label("Email").fill("sales@soundcom.co.za")
    page1.get_by_label("Password").fill("Spaza2314")
    page1.get_by_label("Log In").click()

    print('Spaza login successful!')

    page1.get_by_role('main').get_by_role('link', name="Soundcom Shop").click()
    page1.get_by_role('link', name="Products").click()

    print('Starting stock quantity update...')

    update_stock('ACT311(II)', 8687)
    update_stock('ACT312(II)', 10228)
    update_stock('ACT32H(8B)', 10254)
    update_stock('ACT32T(8B)', 10255)
    update_stock('ACT343', 10229)
    update_stock('BK12V4.5', 10668)
    update_stock('DPM3', 10669)
    update_stock('ELA62', 10252)
    update_stock('MA100', 10209)
    update_stock('MA200', 10222)
    update_stock('MA300D', 10223)
    update_stock('MA707PA', 10225)
    update_stock('MA708B', 10174)
    update_stock('MA808', 10227)
    update_stock('MSP40', 10250)
    update_stock('MU53HN', 10257)
    update_stock('MU53L', 10256)
    update_stock('MRM70(II)(8B)', 10667)
    update_stock('MTG100R', 10735)
    update_stock('MTG100T', 10731)
    update_stock('E20S', 10732)
    update_stock('MU101P', 10737)
    update_stock('MTG100C4', 10734)
    update_stock('MTG100C12', 10735)
    update_stock('MTG100C28', 10736)
    update_stock('PSU', 10670)
    update_stock('SC15RT', 10244)
    update_stock('SC30T', 10245)
    update_stock('SC40T', 10246)
    update_stock('SC420TB', 10241)
    update_stock('SC506F', 10306)
    update_stock('SC530TB', 10242)
    update_stock('SC606F', 10238)
    update_stock('SC610E', 10240)
    update_stock('SC640TB', 10305)
    update_stock('SC80T', 10249)
    update_stock('SC810F', 10239)
    update_stock('SC100', 10671)
    update_stock('SC200', 10672)
    update_stock('SC300', 10673)
    update_stock('SC707', 10674)
    update_stock('SC708', 10675)
    update_stock('SC808', 10676)
    update_stock('T60DTB', 10175)
    update_stock('T120DTB', 10661)
    update_stock('T240DTB', 10662)
    update_stock('T350DTB', 10664)
    update_stock('T500DTB', 10665)
    update_stock('TB40', 10230)
    update_stock('TB60', 10666)
    update_stock('TB120', 10233)

    print('Stock update complete! Closing in 2 minutes...')

    browser1.close()

    timer('End')

    time.sleep(120)

    # https://www.youtube.com/watch?v=wp2pNVUl3lc&t=201s - pyinstaller tutorial  
    # command for compiling to .exe file: pyinstaller main.py --onefile --console