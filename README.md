# 🧾 Automatic Inventory System using Excel VBA

This Excel-based Inventory Management System was developed to help efficiently track and manage stock levels within a company. It provides a structured and semi-automated way to record purchased items, monitor in-garage stock, and manage delivery statuses — all within a familiar Excel interface enhanced by VBA macros.

---

## 🧩 Overview

The system maintains two synchronized datasets:
- **All Records** — a high-level view showing total quantities (including those still in delivery) and in-stock quantities.  
- **Details** — an itemized log that records every stock entry along with serial numbers, supplier details, and sold status.

By integrating both datasets, the workbook provides accurate visibility into which items are **physically available**, **still shipping**, or **already sold**.

---

## ⚙️ Features

✅ **Stock Tracking**
- Record new stock arrivals with detailed item and serial numbers.
- Automatically update in-stock and total quantity values.

✅ **Duplicate Prevention**
- Prevents adding a stock entry where the *item number and serial number* already exist.
- Ensures no duplicate item/serial pairs can be recorded in the same batch or across previous records.

✅ **Quantity Control**
- Guards against adding items that would make the in-stock quantity exceed the total quantity.
- Prevents adding in-stock entries for items that don’t yet exist in the master list (“All Records”).

✅ **Sold Status Automation**
- When a user changes an item’s sold/delivered status, the in-stock count automatically updates and reflects in “All Records”.

✅ **Data Management**
- Users can add, edit, or delete stock records.
- Deleting a record auto-adjusts stock counts (always verify in “All Records” after deletion).

✅ **Search and Filtering**
- A built-in **Search Sheet** allows filtering by item number, serial number, supplier, or sold status.

✅ **User Help**
- Additional manuals and usage notes are available in dedicated “Help” and “Manual” sheets.

---

## 🎬 Demo Video

A short demo video is included in this repository to showcase basic usage of the inventory system.  
Click the video below to view it:

[![Demo Video](https://img.icons8.com/ios-filled/50/000000/video.png)](Inventory%20Management%20Excel%20Demo.mp4)

Or directly access the file in the repo: [Inventory Management Excel Demo.mp4](Inventory%20Management%20Excel%20Demo.mp4)

The video walks through:

- Adding new stock arrivals
- Editing item details
- Using the search sheet to filter items
- Managing sold/delivered status
- Viewing overall records and details

---

## 🚀 How to Use

1. **Enable Macros** when opening the Excel file.  
2. Navigate to the **“+ New Inventories”** sheet to record new stock arrivals.  
3. Fill in item and serial details, then run the `RecordNewStockArrivals` macro (e.g., via a button or Developer tab).  
4. The system will:
   - Validate entries,
   - Check for duplicates,
   - Ensure stock quantities remain within limits,
   - Then record valid entries into the **Details** sheet.
5. To view or modify data:
   - Use **All Records** for quantity overview.
   - Use **Details** for specific item information.
   - Use **Search** to quickly locate an item by number, serial, or supplier.
6. If deleting a record, double-check “All Records” to confirm the quantities update correctly.
7. For additional guidance, refer to the **Help Sheet** or **Demo Video** provided.

---

## 🐞 Known Bug

- The system currently **requires at least one existing record** in both the **Details** and **All Records** sheets.  
- To avoid errors, you should **create a dummy record** if the sheets are empty.  
- The record must be **inside the formatted table**, i.e., within the alternating row colors, for the macros to work correctly.  
- This ensures the system can correctly read and update the stock information.

---

## 🧠 Notes

- Always **save your work** after major updates.  
- Ensure macros are enabled; otherwise, the automation features will not work.  
- For troubleshooting or bug reporting, please contact me via email.

