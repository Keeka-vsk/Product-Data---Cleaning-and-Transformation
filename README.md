# Product-Data---Cleaning-and-Transformation
# Data Cleaning and Transformation in Excel

A hands-on Excel project that takes a messy product dataset and turns it into a clean, analysis-ready table. It covers the core data-preparation skills: handling missing values, fixing inconsistent text, removing duplicates, splitting and merging columns, number formatting, and conditional formatting.

**Tool:** Microsoft Excel  |  **File:** [`Data_Cleaning_and_Transformation_Answers.xlsx`](./Data_Cleaning_and_Transformation_Answers.xlsx)

---

## Final Dataset

After cleaning, the dataset has **31 records** across **6 columns**:

| Column | Description |
|---|---|
| Manufacturing Date | Date of manufacture, displayed as `DD-MM-YYYY` |
| Country Code | Two-letter country code (US, UK, IN, CA, ...) |
| Product Brand | Product and brand combined, e.g. `Laptop-Dell` |
| Price | Unit price, formatted as currency |
| Quantity | Units available |
| Category | Product category (Electronics, Fashion, Kitchen, Accessories, Outdoor) |

**Sample of the cleaned data**

| Manufacturing Date | Country Code | Product Brand | Price | Quantity | Category |
|---|---|---|---|---|---|
| 28-01-2026 | US | Laptop-Dell | 1000.00 | 30 | Electronics |
| 15-02-2026 | US | Sneakers-Nike | 80.00 | 15 | Fashion |
| 03-03-2026 | US | Coffee Maker-Keurig | 130.00 | 40 | Kitchen |
| 11-04-2026 | US | Smartphone-Samsung | 900.00 | 25 | Electronics |
| 22-05-2026 | US | Backpack-North Face | 70.00 | 20 | Accessories |

---

## What I Did

### 1. Handling Missing Values

**Missing prices.** Blank prices were filled with the column mean using a formula, so the fix stays dynamic if the data changes:

```excel
=IF(ISBLANK(D2), AVERAGE($D$2:$D$32), D2)
```

The mean used was **297.86**.

**Missing categories.** Blank categories were filled by matching on product name: if the same product appeared elsewhere with a category, that category was reused.

### 2. Correcting Inconsistent Data

**Product Name.** Inconsistent casing and stray spaces or non-printing characters were standardized with:

```excel
=PROPER(TRIM(CLEAN(B2:B32)))
```

**Category typos.** The typo `Electroni` was replaced with `Electronics` using **Find & Replace** (4 cells fixed).

### 3. Removing Duplicates

Duplicate rows were removed with **Data → Remove Duplicates**, using the Product ID column as the key.

### 4. Splitting and Merging Data

**Split.** The `Product ID` column was split on the `-` delimiter (Data → Text to Columns), producing day, month, and country code. Day and month were then rebuilt into a single **Manufacturing Date** column with `CONCAT`, leaving **Country Code** as its own column.

**Merge.** `Brand Name` and `Product Name` were combined into **Product Brand**:

```excel
=CONCAT(C2, "-", D2)
```

### 5. Number Formatting

- **Price:** formatted as currency via Home → Number → Currency.
- **Manufacturing Date:** displayed as `DD-MM-YYYY`:

```excel
=TEXT(A2, "dd-mm-yyyy")
```

### 6. Conditional Formatting

- **Price:** applied **data bars** to `D2:D32` so relative price is visible at a glance.
- **Category:** created a custom rule highlighting every cell that contains **"Electronics"**:
  Home → Conditional Formatting → New Rule → *Format only cells that contain* → Cell Value → *specific text* → *containing* → `Electronics`.

---

## Skills Demonstrated

- Missing-value imputation (mean fill with `IF` / `ISBLANK` / `AVERAGE`)
- Text cleaning (`PROPER`, `TRIM`, `CLEAN`)
- Find & Replace for typo correction
- De-duplication with Remove Duplicates
- Text to Columns, `CONCAT`, and `TEXT`
- Custom number and date formats
- Conditional formatting (data bars and rule-based highlighting)

---


## How to Use

1. Download `Data_Cleaning_and_Transformation_Answers.xlsx`.
2. Open it in Excel. The cleaned data is in columns **A–F**, and the written answer for each task is in the notes to the right.
3. Click into the cells to inspect the formulas and conditional formatting rules.
