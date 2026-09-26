 📊 Excel Lab 8 – Merging and Retrieving Data Using Lookup Functions

## 🎯 Objective

The objective of this Excel lab is to understand how **lookup functions** can be used to retrieve specific information from a dataset based on a given **Order_ID**.

## 📋 Tasks Completed

| No. | Task | Function Used | Purpose |
|---|---|---|---|
| 1️⃣ | Retrieve Sales for a given Order_ID | **VLOOKUP** | Finds the matching Order_ID and returns its Sales value |
| 2️⃣ | Retrieve Profit for a given Order_ID | **XLOOKUP** | Finds the matching Order_ID and returns its Profit value |
| 3️⃣ | Find City for a given Order_ID | **XLOOKUP** | Finds the matching Order_ID and returns the corresponding City |

## 🧮 Formulas Used

### 🔹 1. VLOOKUP – Retrieve Sales

```excel
=VLOOKUP(Order_ID,Table,Sales_Column,FALSE)

Explanation:
VLOOKUP searches for the specified Order_ID in the first column of the selected table and returns the corresponding Sales value.

🔹 2. XLOOKUP – Retrieve Profit
=XLOOKUP(Order_ID,Order_ID_Column,Profit_Column)

Explanation:
XLOOKUP searches for the given Order_ID in the Order_ID column and returns the corresponding Profit from the Profit column.

🔹 3. XLOOKUP – Find City
=XLOOKUP(Order_ID,Order_ID_Column,City_Column)

Explanation:
XLOOKUP searches for the given Order_ID and returns the corresponding City from the dataset.

🛠️ Functions & Tools Used
Tool / Function	Usage
📗 Microsoft Excel	Data analysis and lookup operations
🔍 VLOOKUP	Retrieve Sales based on Order_ID
⚡ XLOOKUP	Retrieve Profit and City based on Order_ID
📚 Key Learning
Understanding the purpose of lookup functions in Excel.
Retrieving specific values from a large dataset.
Using VLOOKUP for vertical data lookup.
Using XLOOKUP for flexible and efficient data retrieval.
Searching records using a unique Order_ID.
Retrieving Sales, Profit, and City from related data.
✅ Outcome

Successfully performed data retrieval using VLOOKUP and XLOOKUP and obtained the required Sales, Profit, and City information for the given Order_ID.
