import chromedriver_binary
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from bs4 import BeautifulSoup
import pandas as pd
import logging
import warnings
import time
# Suppress all warnings
warnings.filterwarnings('ignore')

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
#open chrome driver
driver = webdriver.Chrome(options=options)
driver.maximize_window()

# Suppressing logging level for selenium and related errors
logging.getLogger('selenium').setLevel(logging.ERROR)


# Specify the path to your Excel file
excel_file_path = r'D:\SEm 8\Results\Batch22CSR.xlsx'
read_df = pd.read_excel(excel_file_path)
rollno = read_df['rollno'].tolist()
invalid_rollno=[]
# rollno = ['20CSR137','20CSR131'] # for sample


#Taking all Semesters available links
allsem = []
if len(rollno)>0 :
    try:
        driver.get('https://results.kongu.edu/allresinrg.php')
        rollno_input_box = driver.find_element(By.XPATH, "//input[@id='regno']")
        rollno_input_box.send_keys(rollno[0])

        click_get_results = driver.find_element(By.XPATH, "//input[@value='Get Results']")
        click_get_results.submit()

        soup= BeautifulSoup(driver.page_source,features="html5lib")
        tables = soup.findAll('table')
        if len(tables) > 1:
            table_body = tables[1].find('tbody')
            if table_body:
                rows = table_body.findAll('tr')
                for i in rows[1:]:
                    link = i.find('a')['href']
                    sem_text = i.find('td').text.strip()
                    exam_text = i.find('a').text.strip()
                    data_dict = {
                        "link": link,
                        "sem": sem_text,
                        "exam": exam_text
                    }
                    allsem.append(data_dict)
            else:
                print("No table body found in the second table.")
        else:
            print("Less than 2 tables found on the page.")
    
        
    except NoSuchElementException:
        print("Element not found on the page. Check if the page structure has changed.")
    except WebDriverException as e:
        print("Selenium WebDriver encountered an exception:", e)
    except :
        print("First Roll No itself invalid - Check it on : First Roll Number must to be valid")

else:
    print("No Roll Numbers Found")

time.sleep(2)
#Marks Extracting
df = pd.DataFrame()
for iter in allsem:
    for j_roll in rollno:
        link = iter["link"].replace(rollno[0],str(j_roll))
        try:    
            driver.get(link)
            time.sleep(0.3)
            soup= BeautifulSoup(driver.page_source,features="html5lib")
            tables = soup.findAll('table')
            #Basic Details
            table_body1 = tables[0].find('tbody').findAll('font')
            name = table_body1[0].text
            rno = table_body1[1].text
            pro_branch = table_body1[2].text
            img_src=tables[0].find('img')
            # print("Name : ",name)
            # print("Roll No : ",rno)
            # print("Dept : ",pro_branch)
            new_row = {'name': name, 'roll_no': rno, 'dept':pro_branch, 'sem':iter["sem"],"exam":iter["exam"],"img_src":"https://results.kongu.edu/"+img_src['src']}
            df = df.append(new_row, ignore_index=True)
            # df = pd.concat([df, new_row], ignore_index=True)
            #Marks  
            table_body2 = tables[1].find('tbody').findAll('tr')
            for i in range(1, len(table_body2)):
                td = table_body2[i].findAll('td')
                th = table_body2[i].findAll('th')
                row_index = df[(df['roll_no'] == rno) & (df['sem']==iter["sem"]) & (df["exam"]==iter["exam"])].index
                df.loc[row_index, "Course_"+str(i)+"_Semester"]=td[0].text
                df.loc[row_index, "Course_"+str(i)+"_Code"]=td[1].text
                df.loc[row_index, "Course_"+str(i)+"_Name"]=td[2].text
                df.loc[row_index, "Course_"+str(i)+"_Credits"]=th[0].get_text(strip=True)
                df.loc[row_index, "Course_"+str(i)+"_Grade"]=th[1].get_text(strip=True)


            #GPA and CGPA
            gpa = tables[2].find('tbody').find('font').text
            cgpa = tables[3].find('tbody').find('font').text
            df.loc[row_index, "GPA"] = gpa.split(':')[1].strip()
            df.loc[row_index, "CGPA"] = cgpa.split(':')[1].strip()
        except:
            invalid_rollno += [j_roll]
        time.sleep(0.25) 
    if invalid_rollno:
        print("Invalid Roll Numbers on ",iter['sem'],"-",iter["exam"])
        print(invalid_rollno)
        invalid_rollno=[]

driver.close()


if not df.empty:
    # Specify the file path and name
    excel_file_path = r'D:\SEm 8\Results\AllSem_Extracted_Results\ESE_All_Sem_22CSE.xlsx'
    # Convert and save the DataFrame to an Excel file
    df.to_excel(excel_file_path, index=False)
    print(f"DataFrame has been saved to {excel_file_path}")

# Reset warnings to default behavior
warnings.resetwarnings()