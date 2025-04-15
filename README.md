# KEC-Grades-Automation

---

## ⚙️ Setup and Usage

1. **Install Dependencies**:
   - Make sure you have the necessary libraries installed. You can install them using `pip`:
     ```bash
     pip install selenium pandas beautifulsoup4 chromedriver-autoinstaller
     ```

2. **Download ChromeDriver**:
   - You need to have **ChromeDriver** installed for Selenium to work. It can be installed via `chromedriver-autoinstaller` or manually.

3. **Prepare Input Data**:
   - The Excel file should contain two columns: `rollno` and `dob`. The script will read this data and use it to extract the student results from the website.
   - Example Excel data:
     | rollno   | dob       |
     |----------|-----------|
     | 20CSR137 | 12.09.2002|
     | 20CSR131 | 20.07.2003|
    - Need to update the results page link in the code.

4. **Run the Script**:
   - Execute the script in the terminal or command prompt:
     ```bash
     python scripts/extract_grades.py
     ```

5. **View Results**:
   - After execution, the results will be saved in an Excel file (e.g., `extracted_results.xlsx`).

---

## 💻 Code Details

### `extract_grades.py`
- This script extracts the grades for the current semester by entering the roll number and date of birth.
- It automates the login process and navigates to the grades page for each student.
- The data is extracted and saved in a DataFrame, which is then exported as an Excel file.

### `extract_all_sem_grades.py`
- This script retrieves the grades for all semesters of a student.
- It requires the roll number and DOB and extracts the results for each available semester by scraping the links from the results page.

---

## 📝 Key Insights

- **Data Extraction**: The grades and academic information are scraped and stored in an Excel file.
- **Error Handling**: Invalid roll numbers or incorrect data will be skipped, and an error log is maintained.
- **Multi-semester Support**: The tool fetches the results for all available semesters and not just the current one.

---

## ❗ Known Issues

- The website structure may change, so the scraping logic may need adjustments if the site layout is modified.
- The scraping speed can be slow due to the need for handling multiple web requests and waiting for pages to load.

---

## 🔧 Requirements

- Python 3.x
- Chrome Browser
- ChromeDriver (matching the version of Chrome)

---

## 📅 Future Enhancements

- **Multi-threading/Parallel Processing**: To speed up the data extraction process.
- **Automated Notifications**: Notify users upon successful completion of data extraction.
- **GUI Interface**: Build a simple graphical user interface for users to easily input roll numbers and view results.

---

## 🏷️ License

This project is open-source and available for personal and educational use.

---
