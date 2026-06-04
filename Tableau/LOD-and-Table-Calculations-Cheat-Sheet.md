# Tableau to Power BI: LOD & Table Calculation Cheat Sheet

## 1. Level of Detail (LOD) Expressions

### FIXED

**Goal:** Calculate a metric at a specific dimension level, completely ignoring whatever filters or dimensions are currently in the visual.

**Tableau:**
`{ FIXED [Region] : SUM([Sales]) }`

**Power BI (DAX):**

```dax
Fixed Region Sales =
CALCULATE(
    SUM(Sales[Sales]),
    ALLEXCEPT(Sales, Sales[Region])
)
```

> **How it works:** `ALLEXCEPT` tells the engine to strip away every filter currently affecting the Sales table _except_ for Region.

### EXCLUDE

**Goal:** Calculate a metric ignoring a specific dimension that is currently in the visual (often used to get a subtotal).

**Tableau:**
`{ EXCLUDE [Category] : SUM([Sales]) }`

**Power BI (DAX):**

```dax
Sales Excluding Category =
CALCULATE(
    SUM(Sales[Sales]),
    REMOVEFILTERS(Products[Category])
)
```

> **How it works:** `REMOVEFILTERS()` (or `ALL()`) specifically deletes the filter context coming from the Category column, acting exactly like an EXCLUDE.

### INCLUDE

**Goal:** Calculate a metric at a lower level of granularity than what is shown in the visual, then aggregate it back up (e.g., average customer lifetime value by region).

**Tableau:**
`AVG({ INCLUDE [Customer ID] : SUM([Sales]) })`

**Power BI (DAX):**

```dax
Avg Sales per Customer =
AVERAGEX(
    VALUES(Customers[CustomerID]),
    CALCULATE(SUM(Sales[Sales]))
)
```

> **How it works:** DAX uses an iterator function (`AVERAGEX`). `VALUES` creates a virtual table of all unique customers in the current context. DAX calculates the total sales for each customer row by row, then averages the results.

---

## 2. Table Calculations

### Percent of Total

**Goal:** What percentage does this row contribute to the grand total?

**Tableau:**
`SUM([Sales]) / TOTAL(SUM([Sales]))`

**Power BI (DAX):**

```dax
% of Total Sales =
DIVIDE(
    SUM(Sales[Sales]),
    CALCULATE(SUM(Sales[Sales]), REMOVEFILTERS(Sales))
)
```

> **How it works:** You divide the current row's sales by a `CALCULATE` statement that uses `REMOVEFILTERS(Sales)` to get the grand total of the entire table.

### Running Total (YTD)

**Goal:** A cumulative sum over time.

**Tableau:**
`RUNNING_SUM(SUM([Sales]))`

**Power BI (DAX):**

```dax
YTD Sales =
TOTALYTD(
    SUM(Sales[Sales]),
    'Date'[Date]
)
```

> **How it works:** DAX has native Time Intelligence functions, provided you have a continuous Date table. `TOTALYTD` handles the accumulation automatically.

### Difference From Previous (e.g., Month-over-Month)

**Goal:** Subtract last month's value from this month's value.

**Tableau:**
`ZN(SUM([Sales])) - LOOKUP(ZN(SUM([Sales])), -1)`

**Power BI (DAX):**

```dax
MoM Sales Diff =
VAR CurrentMonth = SUM(Sales[Sales])
VAR PreviousMonth =
    CALCULATE(
        SUM(Sales[Sales]),
        DATEADD('Date'[Date], -1, MONTH)
    )
RETURN
    CurrentMonth - PreviousMonth
```

> **How it works:** Instead of relying on a visual `LOOKUP`, DAX uses `DATEADD` to physically shift the filter context back by one month on the Date table. (Using `VAR` makes DAX much easier to read and debug).

### Rank

**Goal:** Rank items based on a metric.

**Tableau:**
`RANK(SUM([Sales]))`

**Power BI (DAX):**

```dax
Rank by Sales =
RANKX(
    ALL(Products[ProductName]),
    CALCULATE(SUM(Sales[Sales])),
    ,
    DESC
)
```

> **How it works:** `RANKX` iterates through a table. `ALL()` ensures you are ranking the product against _all_ products, not just the one currently visible in the row context.
