# CRM Log Task Automation System

This is a Python automation program that helps log CRM tasks automatically using phone numbers and remarks from Excel.

---

## Features

- Reads phone numbers and remarks from Excel  
- Automatically searches student in CRM  
- Handles both Lead and Opportunity records  
- Creates new task under "Related"  
- Auto fills:
  - Comments  
  - Type (Contact)  
  - Sub Type (Outbound Call)  
  - Subject (Outbound Call)  
  - Due Date (Today)  
  - Status (based on remarks)  
- Automatically clicks Save  
- Processes multiple records continuously  

---

## Technologies

- Python  
- Selenium  
- Pandas  
- Chrome WebDriver  

---

## Purpose

This project was created as a real-world automation tool to reduce repetitive manual CRM work.  

Instead of manually searching, clicking, and filling forms for every student, this system performs everything automatically.

It also helped me improve my understanding of:
- Web automation  
- Debugging real UI issues  
- Handling dynamic web elements (Vue / Element UI)  
- Writing more stable and reliable scripts  

---

## Challenges Faced

While developing this system, I encountered several real-world issues:

- Elements not detected even though visible on screen  
- Clicking wrong buttons (e.g. "Edit" instead of student name)  
- Multiple results returned for a single phone number  
- Dynamic dropdown issues (especially Status selection)  
- Page loading delays causing script failure  
- "invalid element state" and stale element errors  
- Multiple "Status" fields leading to incorrect selection  

---

## Solutions

To overcome these challenges, I improved the script by:

- Using more precise XPath targeting  
- Filtering correct elements and excluding irrelevant ones  
- Implementing proper waiting strategies (WebDriverWait)  
- Resetting the page before each loop to avoid stale states  
- Targeting only visible dropdown elements  
- Using JavaScript click for elements that cannot be clicked normally  
- Adding fallback and retry logic for better stability  

---

## Result

The system is now stable and able to:

- Run continuously without crashing  
- Handle different edge cases  
- Reduce manual workload significantly  

---

## Author

Bryan Lim  

---

## Code

See full implementation here:  
