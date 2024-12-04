# Excel

## Dataset

The dataset is stored in the Excel file `[Product Sales Data.xlsx](https://docs.google.com/spreadsheets/d/1PKOLL6dkTa7aMKcNca1ee6_DXKlM1pKyhhpV1N-aBpE/edit?usp=sharing)` and contains the following columns:

- **Product Name**: Name of the product.
- **Brand Name**: Brand associated with the product.
- **Price**: Price of the product.
- **Category**: Product category (e.g., Electronics, Clothing, etc.).

## Project Overview

### Assignment 1: Excel Table Creation and Formatting
- Created a table to organize product information, including the columns: 'Product Name,' 'Brand Name,' 'Price,' and 'Category.'
- Formatted the table with different designs to enhance readability and analysis.

### Assignment 2: Excel Formulas and Conditional Functions
- **Sum, Count, Average**:
  - Calculated the total price of all products in the dataset using the `SUM` function.
  - Counted the total number of products using the `COUNT` function.
  - Calculated the average price of products using the `AVERAGE` function.
- **Min and Max**:
  - Determined the minimum price using the `MIN` function.
  - Determined the maximum price using the `MAX` function.
- **SUMIF and COUNTIF**:
  - Calculated the total price for products in the 'Electronics' category using the `SUMIF` formula.
  - Determined the count of products with a price greater than $20 using the `COUNTIF` formula.
- **Text Formatting**:
  - Used the `LEFT` function to extract the first 3 characters of the 'Product Name.'
  - Used the `RIGHT` function to extract the last 5 characters from the 'Brand Name' column.
  - Used the `MID` function to extract characters 2 to 5 from the 'Category' column.
- **IF Statements**:
  - Categorized products with a price greater than $25 as 'High Price' and others as 'Standard Price' using the `IF` function.

### Assignment 3: Data Cleaning and Handling Missing Data
- **Removing Duplicates**:
  - Identified and removed any duplicate rows from the dataset.
- **Handling Empty Rows and Columns**:
  - Removed any completely empty rows and redundant columns from the dataset.
- **Dealing with Missing Values**:
  - Checked for missing values in the 'Price' column and proposed strategies to handle missing values. (e.g., imputation or removal).

### Assignment 4: Data Wrangling
- **Data Type Conversion**:
  - Checked the data types of the 'Price' and 'Quantity' columns and converted them to the appropriate data types (numeric).
- **Transforming Text Data**:
  - Examined and cleaned the 'Product Name' and 'Brand Name' columns for inconsistencies or irregularities, ensuring data consistency.

 ## Data (https://docs.google.com/spreadsheets/d/16aPT1v4eVY7VhplV47_FsPCfKLbHPNZzs7og0vetzIQ/edit?usp=drive_link)
### Assignment 5: Filtering, Sorting, and Lookup Functions
- **Filtering by Price**:
  - Filtered the dataset to display only products with a price greater than $25.
- **Category-based Filtering**:
  - Created a filtered view to show only products in the 'Electronics' category using Excel’s filtering features.
- **Sorting by Price**:
  - Sorted the dataset in ascending order based on the 'Price' column.
- **Multilevel Sorting**:
  - Implemented a multilevel sort to first sort by 'Category' in ascending order and then by 'Price' in descending order within each category.
- **VLOOKUP for Category Information**:
  - Used the `VLOOKUP` function to retrieve the 'Category' information for a specific product based on its 'Product Name.'
- **HLOOKUP for Price Information**:
  - Used the `HLOOKUP` function to find the 'Price' of a product based on its 'Product Name.'

### Assignment 6: Pivot Tables and Trend Analysis
- **Basic Pivot Table Creation**:
  - Created a pivot table to summarize the total sales for each 'Category' in the dataset. Applied a filter to show only 'Electronics' and 'Clothing' categories.
- **Bar Chart for Category Sales**:
  - Created a bar chart to visually represent the total sales for each category.
- **Pie Chart for Brand Distribution**:
  - Generated a pie chart to illustrate the distribution of products across different brands.
- **Scatter Plot for Price vs. Quantity**:
  - Constructed a scatter plot to explore the relationship between 'Price' and 'Quantity' for each product in the dataset.

## Technologies Used

- Microsoft Excel
- Excel Formulas (SUM, AVERAGE, COUNT, IF, etc.)
- Pivot Tables and Charts
- Data Preprocessing and Analysis

