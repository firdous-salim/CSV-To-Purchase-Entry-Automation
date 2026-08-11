# 📊 CSV to Automated Purchase Entry System

A robust and high-speed Windows Forms (VB.NET) desktop application designed to automate bulk purchase data entry. It processes complex supplier invoices (100-150 bills) in under a minute, mapping CSV data directly into a relational MS Access database with complete accuracy.

## 📸 Screenshots
<img width="1366" height="768" alt="image" src="https://github.com/user-attachments/assets/fd168c96-bc57-4fcb-bc0f-2ac234769d06" />

---

## 🚀 Key Features

### ⚙️ 1. Dynamic Supplier Formatting (One-Time Setup)
* Suppliers often use varying CSV structures. This system allows users to create and save custom column-mapping templates for different suppliers.
* Once saved, users can simply select the supplier, and the system automatically formats the incoming CSV to match internal database schemas.

### 🧮 2. Live Preview & Real-Time Calculations
* Loads CSV data into a DataGridView for a visual preview before database insertion.
* **Automated Computations:** Instantly calculates Deal Adjustments, Discounts, GST, Free Items, and Net Rate directly within the grid.

### 🧠 3. "Smart Mapper" Product Resolution
* Scans incoming CSV data against the internal `Product Master` to assign product IDs automatically.
* If an exact match is not found, the custom **Smart Mapper** algorithm displays similar existing products, allowing the user to manually select and resolve discrepancies.
* Capable of parsing and handling complex, multiplied data formats (e.g., parsing "1226").

### 🛡️ 4. ACID-Compliant Database Transactions
* **Zero Duplicate Entries:** System architecture prevents duplicate billing entries.
* **Fail-Safe Processing:** Uses strict SQL transactions when inserting data across multiple tables (`Batch Master`, `Purchase Master`, `Purchase Details`, `Tax`, and 3 additional normalized tables). 
* If a single error occurs during the bulk insert, the entire transaction is rolled back, ensuring absolute data integrity.

---

## 💻 Tech Stack
* **Language/Framework:** VB.NET, Windows Forms (WinForms) .NET Framework 4.5
* **Database:** MS Access (Optimized for offline desktop utilization)
* **Architecture:** Multi-table relational database with strict transactional scopes.

## ⚙️ How to Run This Project
1. Clone this repository to your local machine.
2. Open the `.sln` file using Visual Studio.
3. Ensure MS Access Database Engine is installed on your system.
4. Update the DB connection string in the configuration file to point to your local `.accdb` or `.mdb` file.
5. Press F5 to build and run the application.
