# FUNDAMETAL-BOOSTER-EXCEL

# 🚀 PR. 1 Fundamental Booster

## 🔍 Overview

**PR. 1 Fundamental Booster** is an Excel practice project that builds a strong foundation in formulas and functions. Each worksheet applies a specific concept to realistic datasets such as **student scores, product sales, and employee records**.

The project demonstrates how to **classify, count, look up, clean, and extract data** efficiently using Excel's built-in functions.

---

## 🧰 Topics Covered

| Category | Functions / Concepts |
|----------|----------------------|
| 📌 References & Formatting | Relative (`A1`) & Absolute (`$A$1`) references, data input, custom formats |
| 🔀 Logical | `IF`, Nested `IF`, `IF + AND`, `IF + OR` |
| 🔢 Conditional Aggregation | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` |
| 🔎 Lookup | `VLOOKUP`, `XLOOKUP`, `XMATCH` |
| 🎯 Dynamic Search | `INDEX` + `MATCH` |
| ✂️ Text | `LEFT`, `FIND`, `UPPER`, `LOWER` |
| 🔗 Dynamic References | `INDIRECT`, `OFFSET` |
| 📅 Date & Time | Age calculation, date difference |
| ➗ Math | `ROUND`, `CEILING`, `FLOOR` |
| 🧲 Multi-value Output | `FILTER` |

---

## ✅ Main Tasks

| # | Task | Function(s) |
|---|------|-------------|
| 1 | Classify student grades | `IF` |
| 2 | Count students scoring above 60 | `COUNTIFS` |
| 3 | Find a product price from sales data | `VLOOKUP` |
| 4 | Fetch employee salaries dynamically | `XLOOKUP` |
| 5 | Format and analyze date of joining | Date functions |
| 6 | Extract students with grades above 80 | `FILTER` |

---

## 🧮 Formulas & Solutions

> Cell references below are examples. Adjust them to match your worksheet layout.

### 1️⃣ Relative & Absolute References, Formatting
```excel
=B2*$F$1          ' $F$1 stays fixed (e.g., tax rate) while B2 changes
```
Formatting applied: **bold, italic, currency, date, and custom number formats**.

### 2️⃣ IF & Nested IF
```excel
' Grade classification
=IF(B2>=90,"A+",IF(B2>=80,"A",IF(B2>=70,"B",IF(B2>=60,"C",IF(B2>=50,"D","Fail")))))

' Discount based on price threshold
=IF(C2>=1000,C2*0.20,IF(C2>=500,C2*0.10,0))
```

### 3️⃣ IF with AND / OR
```excel
' Scored above 80 in both Math & Science
=IF(AND(B2>80,C2>80),"Distinction","Not Eligible")

' Product eligible for discount
=IF(OR(D2="Electronics",E2>500),"Eligible","Not Eligible")
```

### 4️⃣ COUNTIFS, SUMIFS, AVERAGEIFS
```excel
' Students scoring above 50 in Math
=COUNTIFS(B2:B50,">50")

' Total sales for a region and product
=SUMIFS(D2:D100,A2:A100,"North",B2:B100,"Laptop")

' Average score of students above 60
=AVERAGEIFS(B2:B50,B2:B50,">60")
```

### 5️⃣ VLOOKUP
```excel
' Student name from ID
=VLOOKUP(A2,Students!$A$2:$B$50,2,FALSE)

' Product price from product code
=VLOOKUP(A2,Products!$A$2:$C$30,3,FALSE)
```

### 6️⃣ INDEX & MATCH
```excel
' Sales of a salesperson in a specific month
=INDEX(B2:M20,MATCH("Rahul",A2:A20,0),MATCH("March",B1:M1,0))

' Dynamic employee detail lookup (no column restriction)
=INDEX(A2:F50,MATCH(H1,A2:A50,0),MATCH(H2,A1:F1,0))
```

### 7️⃣ TEXT Functions
```excel
' First name from full name
=LEFT(A2,FIND(" ",A2)-1)

' Change case
=UPPER(A2)
=LOWER(A2)
```

### 8️⃣ XLOOKUP
```excel
' Flexible salesperson search
=XLOOKUP(H1,A2:A50,D2:D50,"Not Found")

' Employee salary (no sorting needed)
=XLOOKUP(H1,A2:A50,E2:E50,"Not Found")
```

### 9️⃣ XMATCH
```excel
' Position of a product in the sales list
=XMATCH("Laptop",A2:A50)
```

### 🔟 INDIRECT & OFFSET
```excel
' Dynamically reference a range using a cell value
=SUM(INDIRECT(H1&"!B2:B20"))

' Dynamic range for sales trend
=SUM(OFFSET(B2,0,0,H1,1))
```

### 1️⃣1️⃣ Date & Time Functions
```excel
' Age from date of birth
=DATEDIF(B2,TODAY(),"Y")

' Difference in days between two dates
=C2-B2
=DAYS(C2,B2)
```

### 1️⃣2️⃣ Math Functions
```excel
=ROUND(A2,2)
=CEILING(A2,5)
=FLOOR(A2,5)
```

### 1️⃣3️⃣ FILTER
```excel
' Students with grades above 80
=FILTER(A2:C50,C2:C50>80,"No results")

' Top performers (sorted)
=SORT(FILTER(A2:C50,C2:C50>80),3,-1)
```

---

## 📂 Project Structure

```text
PR1-Fundamental-Booster/
│
├── 📄 PR1_Fundamental_Booster.xlsx   # Workbook with all worksheets
├── 📄 README.md                      # Project documentation
└── 📁 images/                        # Screenshots for the README
```

---
