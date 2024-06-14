if __name__ == "__main__":
    from selenium import webdriver
    from parsel import Selector 
    import time 
    import pyodbc
    from time import sleep
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait 
    from selenium.webdriver.chrome.options import Options 
    import os
    import re
    import pandas as pd
    import datetime
    from openpyxl.workbook import Workbook
    from datetime import datetime, timedelta
    from datetime import datetime
    
    parent_path = os.getcwd()
    opts= Options()  
    opts.add_argument('--start-maximized')
    url="https://www.liebertpub.com/doi/full/10.1089/thy.2023.29156.abstracts"
    session_id=0
    start_row=1
    driver = webdriver.Chrome(options=opts)
    driver.maximize_window
    

def convert_to_normal_date(datetime_str):
    # Parse the input datetime string
    datetime_obj = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')

    # Format the datetime as 'Month Day, Year'
    formatted_date = datetime_obj.strftime('%B %d, %Y')

    return formatted_date


def convert_to_standard_date(date_str, year):
    # Parse the input date string and add the year
    date_parts = date_str.split(' ')
    month_str, day_str = date_parts[0], date_parts[1]
    month = datetime.strptime(month_str, '%b').month
    day = int(day_str)

    # Create a datetime object with the provided year
    standard_date = datetime(year, month, day)

    # Format the standard date as 'YYYY-MM-DD'
    formatted_date = standard_date.strftime('%Y-%m-%d')

    return formatted_date




driver.get('https://clubwyndham.wyndhamdestinations.com/us/en/owner-guide/resources/reservations/owner-priority-reservations')
WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH,'//div[@id="onetrust-close-btn-container"]'))).click()

# sleep(2) 
dataList=[]
try:
    for yr in driver.find_elements(By.XPATH,'//div[@class="cell small-12 Auto large-6"]'):
        yr_txt = yr.find_element(By.XPATH,'.//div[@class="content-slice-text-caption-color-default dynamic-content-slice-title-default"]').text
        # print(yr_txt)
        for clel in yr.find_elements(By.XPATH,'.//a[@class="gaChecker"][@data-eventcategory="accordion"]'):

            # if True:
            try:
                clubname = clel.find_element(By.XPATH,'.//div[@class="accordionV3-title accordion-title accordionV3-title-body-1 accordionV3-title-color-primary"]').text
                print(clubname)
                clel.click()
                sleep(1)

                dates = clel.find_element(By.XPATH,'.//following-sibling::*//div[@class="dynamicContentSlice scrollMargin grid-x content-slice-border-default       desktop-no-margin-bottom desktop-no-margin-top desktop-no-padding-bottom desktop-no-padding-top tablet-no-margin-bottom tablet-no-margin-top tablet-no-padding-bottom tablet-no-padding-top no-margin-bottom no-margin-top no-padding-bottom no-padding-top "]|.//following-sibling::*//div[@class="dynamicContentSlice scrollMargin grid-x content-slice-border-default       desktop-no-margin-bottom desktop-no-margin-top desktop-no-padding-bottom desktop-no-padding-top tablet-no-margin-bottom tablet-no-margin-top tablet-no-padding-bottom tablet-no-padding-top no-margin-bottom no-margin-top no-padding-bottom no-padding-top "]/p').text
                # print(dates)
                # print(">>",type(dates))
                
                print("=="*50)
                allDates = dates.split(":")
                # print("!!!$",allDates)
                for i,data in enumerate(allDates[1:]):
                    data1 = data.split('\n')[0]
                    # print(">>>",data1)
                    start_date = data1.replace(".","").replace("  "," ").split("-")[0].strip()
                    month = start_date.strip().split(" ")[0].strip()
                    # print("start",start_date)
                    end_date = data1.replace(".","").split("-")[1].split(",")[0].strip()
                    if len(end_date)<=2:
                        end_date = f"{month} {end_date}"
                    # print("end",end_date)
                    year = data1.split(",")[1].split()[0].strip()
                    # print(f"year:-{year}")
                

                    # Example usage:
                    dateStart = start_date
                    year = year
                    formatted_date1 = convert_to_standard_date(str(dateStart), int(year))
                    # print(formatted_date1)
                    # # Example usage:  
                    dateEnd = end_date
                    formatted_date2 = convert_to_standard_date(str(dateEnd),int(year))
                    # print(formatted_date2)
                    date_list = []

                    formatted_date1 = datetime.strptime(formatted_date1, '%Y-%m-%d')
                    formatted_date2 = datetime.strptime(formatted_date2, '%Y-%m-%d')
                    while formatted_date1 <= formatted_date2:
                        print(convert_to_normal_date(str(formatted_date1)))
                    #     formatted_date1

                        date_list.append(convert_to_normal_date(str(formatted_date1)))
                        formatted_date1 += timedelta(days=1)

                    dataList.append([yr_txt,clubname,date_list])
            except:
                pass   
except Exception as e:
    print("Xpath Not Found")
sleep(0.5)
# sleep(2)
curelem = yr.find_element(By.XPATH,'.//div[@class="content-slice-text-caption-color-default dynamic-content-slice-title-default"]')
driver.execute_script("arguments[0].scrollIntoView();", curelem)
    # sleep(1000)

newList = []
current_run_date = datetime.now().strftime("%Y-%m-%d")

for tempData in dataList:
    for i in tempData[2]:
        newList.append([tempData[0], tempData[1], i, current_run_date,"Automation"])
df = pd.DataFrame(newList, columns=["Year", "Resort", "BlackoutDates", "RunDate","AddedBy"])

df.to_excel("data111.xlsx", index=False)



# sleep(2)
# def insert_data_into_database(year, resort, blockoutdate, rundate):
#     try:
#         server_name = "firstsqlconnection"
#         database_name = "sa"
#         username = "admin"
#         password = "1234"

#         connection_string = f"DRIVER={{SQL Server}};SERVER={server_name};DATABASE={database_name};UID={username};PWD={password}"

#         # Establish the database connection
#         conn = pyodbc.connect(connection_string)

#         # Create a cursor object to interact with the database
#         cursor = conn.cursor()

#         # Define the SQL INSERT statement
#         sql_insert = "INSERT INTO YourTableName (year, resort, blockoutdate, rundate) VALUES (?, ?, ?, ?)"

#         # Execute the SQL INSERT statement
#         cursor.execute(sql_insert, (year, resort, blockoutdate, rundate))

#         # Commit the changes to the database
#         conn.commit()

#         # Close the cursor and the database connection
#         cursor.close()
#         conn.close()

#         print("Data inserted successfully into the database.")
#     except Exception as e:
#         print(f"An error occurred while inserting data into the database: {str(e)}")





# current_run_date = datetime.now().strftime("%Y-%m-%d")


# # Inside your loop
# for tempData in dataList:
#     for i in tempData[2]:
#         year = tempData[0]
#         resort = tempData[1]
#         blackoutdate = i
#         rundate = current_run_date

#         # Call the function to insert data into the database
#         insert_data_into_database(year, resort, blackoutdate, rundate)













