# Excel Project Documentation - Coffee Sales
Analyzing Coffee Sales data in Excel. <br>
Author: Leonardo Beduschi Iunes <br>
Date: 22/01/2025

## Introduction
This project was inspired by and developed using data and ideas from the [The ONLY EXCEL PORTFOLIO PROJECT YOU NEED](https://www.youtube.com/watch?v=m13o5aqeCbM). I would like to acknowledge the video creator for providing valuable insights and resources that helped me complete this project.

## Project Overview

This project involves using Excel to clean and analyze data related to coffee sales. The analysis focused on various aspects of the dataset, including total sales over time, sales by country, the top 5 customers, and other key metrics. Advanced Excel features such as pivot tables, charts, and formulas were utilized to extract insights and present the data in a meaningful way.


## Steps Involved
### **1. Data Import**  
   The dataset was imported from a CSV file provided in the video. <br>
   Link: [Data](https://github.com/mochen862/excel-project-coffee-sales/blob/main/coffeeOrdersData.xlsx)<br>
   <img src="https://github.com/user-attachments/assets/756fd9fc-ce27-4a21-80d1-ce5ca010d63b" alt="Image" width="500"/>

   
### **2. Data Cleaning**  
   

During the data cleaning phase, the following steps were undertaken to complete and clean the data:

-  **XLOOKUP**  
   Used the `XLOOKUP` function to retrieve specific data points from other tables or ranges, ensuring accurate and consistent information.

-  **INDEX MATCH**  
   Utilized the `INDEX` and `MATCH` functions together to perform advanced lookups when `XLOOKUP` was not applicable, offering flexibility in retrieving data.

-  **Multiplication Formula for Sales**  
   Created a formula to calculate total sales by multiplying units sold by the price per unit, ensuring accurate sales figures for analysis.

-  **Multiple IF Functions**  
   Applied multiple `IF` functions to categorize data and handle complex conditions for better organization and insights.

-  **Date Formatting**  
   Standardized date formats to ensure consistency across the dataset, simplifying sorting and filtering tasks.

-  **Number Formatting**  
   Adjusted the number formatting to display values in the correct format (e.g., currency, percentages), improving readability.

-  **Check for Duplicates**  
   Identified and removed duplicate records to ensure the dataset's integrity and accuracy.

-  **Convert Range to Table**  
   Converted data ranges into structured Excel tables to leverage dynamic table features like sorting, filtering, and structured references.

These steps ensured the dataset was clean, complete, and ready for analysis.

<img src="https://github.com/user-attachments/assets/8b4be4bb-5ac7-4feb-87ab-df041c6939d1" alt="Image" width="800"/>

### **3. Analysis Techniques**  
   During the analysis phase, the following techniques were employed:

-  **Pivot Tables**  
   - Constructed pivot tables to summarize and analyze data, facilitating the identification of patterns and trends.
   - Utilized pivot tables to compute aggregates such as sums, averages, and counts across various dimensions.

-  **Updating the Pivot Table Data Source**  
   - Ensured that pivot tables were linked to dynamic data sources, allowing for automatic updates when new data was added.
   - Maintained data integrity by regularly refreshing pivot tables to reflect the most current information.

### **4. Visualization**  
   To enhance data interpretation and presentation, the following visualization processes were implemented:

-  **Pivot Charts**  
   - Developed pivot charts based on pivot table data to provide graphical representations of analytical findings.
   - Applied formatting options to improve readability and highlight key insights within the charts.

-  **Inserting Timelines**  
   - Added timeline controls to filter data by specific time periods, enabling temporal analysis.
   - Customized the appearance of timelines to align with the overall dashboard design.

-  **Inserting Slicers**  
   - Implemented slicers to allow interactive filtering of data across different categories, enhancing user engagement.
   - Formatted slicers for consistency and ease of use within the dashboard.

-  **Building the Dashboard**  
   - Integrated pivot tables, pivot charts, timelines, and slicers into a cohesive dashboard.
   - Focused on creating an intuitive layout that facilitates quick insights and decision-making.

## Dashboard
![Image](https://github.com/user-attachments/assets/b7c86668-64cf-42b0-8047-9ec6f536fff3)
<br>
**Interactive Dashboard:**
[coffeeOrdersData.xlsx](https://github.com/user-attachments/files/18510079/coffeeOrdersData.xlsx)

## Conclusion

This project provided valuable hands-on experience in cleaning, analyzing, and visualizing data using Excel. Through this process, I improved my skills in working with advanced Excel features such as pivot tables, charts, XLOOKUP, and dynamic dashboards. Additionally, I developed a better understanding of how to present data in a clear and actionable way. This project not only enhanced my technical abilities but also deepened my analytical thinking, which will be beneficial for future data-related tasks.
