import chromedriver_binary
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
#open chrome driver
driver = webdriver.Chrome(options=options)
driver.maximize_window()

df = pd.DataFrame()
# Specify the path to your Excel file
excel_file_path = r'D:\SEm 8\Results\Batch22CSR.xlsx'
read_df = pd.read_excel(excel_file_path)
rollno = read_df['rollno'].tolist()
dob = read_df['dob'].tolist()
invalid_rollno=[]
# rollno = ['20CSR137','20CSR131']
# dob = ['12.09.2002','20.07.2003']

for k in range(len(rollno)):
    try:    
        driver.get('https://results.kongu.edu/xxiir/')
        rollno_input_box = driver.find_element(By.XPATH, "//input[@id='regno']")
        rollno_input_box.send_keys(rollno[k])

        dob_input_box = driver.find_element(By.XPATH, "//input[@id='input-popup']")
        dob_input_box.send_keys(dob[k])

        click_get_results = driver.find_element(By.XPATH, "//input[@value='Get Results']")
        click_get_results.submit()

        soup= BeautifulSoup(driver.page_source)

        tables = soup.findAll('table')

        #Basic Details
        table_body1 = tables[0].find('tbody').findAll('font')
        name = table_body1[0].text
        rno = table_body1[1].text
        pro_branch = table_body1[2].text
        # print("Name : ",name)
        # print("Roll No : ",rno)
        # print("Dept : ",pro_branch)
        new_row = {'name': name, 'roll_no': rno, 'dept':pro_branch}
        df = df.append(new_row, ignore_index=True)

        #Marks
        table_body2 = tables[1].find('tbody').findAll('tr')
        for i in range(1, len(table_body2)):
            td = table_body2[i].findAll('td')
            th = table_body2[i].findAll('th')
            row_index = df[df['roll_no'] == rno].index
            df.loc[row_index, td[2].text+" Semester"] = td[0].text
            df.loc[row_index, td[2].text+" Course_Code"] = td[1].text
            df.loc[row_index, td[2].text+" Credits"]= th[0].get_text(strip=True)
            df.loc[row_index, td[2].text+" Grade"] = th[1].get_text(strip=True)

        #GPA and CGPA
        table_body3 = tables[2].find('tbody').findAll('tr')
        if len(table_body3)>2:
            gpa_score = (table_body3[0].find('font')).find('font').text
            cgpa_score = (table_body3[1].find('font')).find('font').text
            # print("GPA : ",gpa_score)
            # print("CGPA : ",cgpa_score)
            df.loc[row_index, "GPA"] = gpa_score[:-1]
            df.loc[row_index, "CGPA"] = cgpa_score[:-1]
        else:
            df.loc[row_index, "GPA"] = "NA"
            df.loc[row_index, "CGPA"] = "NA"
    except:
        invalid_rollno += [rollno[k]]
        # print(rollno[k], " Invalid Credentials")
driver.close()
print("Invalid Credentials : ", invalid_rollno)
# Specify the file path and name
excel_file_path = r'D:\SEm 8\Results\ESE_Jan2024_Batch22CSR_CSE.xlsx'

# Convert and save the DataFrame to an Excel file
df.to_excel(excel_file_path, index=False)


print(f"DataFrame has been saved to {excel_file_path}")