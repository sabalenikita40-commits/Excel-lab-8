📊 Excel Lab 8 – Merging and Retrieving Data Using Lookup Functions

## 🎯 Objective

The objective of this Excel lab is to understand and apply different lookup functions to retrieve specific information from a dataset using a given **Order_ID**.

## 📋 Tasks Completed

| No. | Task | Function Used | Purpose |
|---|---|---|---|
| 1️⃣ | Retrieve Sales for a given Order_ID | **VLOOKUP** | Retrieves the Sales value for the matching Order_ID |
| 2️⃣ | Retrieve Profit for a given Order_ID | **XLOOKUP** | Retrieves the Profit value for the matching Order_ID |
| 3️⃣ | Find City for a given Order_ID | **INDEX + MATCH** | Finds and returns the City corresponding to the Order_ID |

## 🧮 Formulas Used

### 🔹 1. VLOOKUP – Retrieve Sales

```excel
=VLOOKUP(Order_ID,Table,Sales_Column,FALSE)

Explanation:
VLOOKUP searches for the specified Order_ID in the dataset and returns the corresponding Sales value.

🔹 2. XLOOKUP – Retrieve Profit
=XLOOKUP(Order_ID,Order_ID_Column,Profit_Column)

Explanation:
XLOOKUP searches for the given Order_ID in the Order_ID column and returns the corresponding Profit value.

🔹 3. INDEX + MATCH – Find City
=INDEX(City_Column,MATCH(Order_ID,Order_ID_Column,0))

Explanation:
The MATCH function finds the position of the given Order_ID, while INDEX returns the City from the corresponding position.

🛠️ Tools & Functions Used
Tool / Function	Usage
📗 Microsoft Excel	Data analysis and retrieval
🔍 VLOOKUP	Retrieve Sales
⚡ XLOOKUP	Retrieve Profit
🔎 INDEX + MATCH	Find City
📚 Key Learning
Understanding different lookup functions in Excel.
Using VLOOKUP to retrieve data vertically.
Using XLOOKUP for flexible data retrieval.
Combining INDEX and MATCH to retrieve values based on a matching condition.
Searching records using a unique Order_ID.
Retrieving Sales, Profit, and City efficiently from the dataset.
✅ Outcome

Successfully performed data retrieval using VLOOKUP, XLOOKUP, and INDEX + MATCH functions to obtain the required Sales, Profit, and City details based on the given Order_ID.
