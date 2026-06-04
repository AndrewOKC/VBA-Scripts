# Tableau to Power BI: Conditional Aggregations Cheat Sheet

In Tableau, embedding row-level `IF` statements directly inside aggregations is second nature. In Power BI, doing row-by-row evaluation is heavily frowned upon for performance reasons. Instead, DAX requires you to think about **filtering sets of data**. You apply a filter to the data model first, and _then_ run the aggregation.

## 1. The Direct Translation: `CALCULATE` is Your IF Statement

To do conditional math in DAX, you turn to the `CALCULATE()` function. It evaluates an expression (your math) in a modified filter context (your `IF` condition).

**Tableau:**
`SUM(IF [Category] = 'Tech' THEN [Sales] END)`

**Power BI (DAX):**

```dax
Tech Sales =
CALCULATE(
    SUM(Sales[Sales]),
    Products[Category] = "Tech"
)
```

> **How it works:** You are telling DAX: "Take the total sum of sales, but before you do the math, filter the `Products` table down to only show 'Tech'." Power BI's VertiPaq engine handles this columnar filtering instantly.

---

## 2. Handling Multiple Conditions (AND / OR)

Tableau handles complex logic with standard `AND`/`OR` text operators. DAX handles this by either adding multiple filter arguments to `CALCULATE` or using DAX operators (`&&` for AND, `||` for OR).

### Scenario A: The "AND" Condition

**Tableau:**
`SUM(IF [Category] = 'Tech' AND [Region] = 'West' THEN [Sales] END)`

**Power BI (DAX):**

```dax
Tech Sales West =
CALCULATE(
    SUM(Sales[Sales]),
    Products[Category] = "Tech",
    Geography[Region] = "West"
)
```

> **Note:** Every new line separated by a comma inside `CALCULATE` automatically acts as an `AND` condition.

### Scenario B: The "OR" Condition

**Tableau:**
`SUM(IF [Category] = 'Tech' OR [Category] = 'Furniture' THEN [Sales] END)`

**Power BI (DAX):**

```dax
Tech or Furniture Sales =
CALCULATE(
    SUM(Sales[Sales]),
    Products[Category] IN {"Tech", "Furniture"}
)
```

> **Note:** While you can use the `||` operator, using the `IN` syntax with curly brackets `{}` is vastly cleaner and easier to read when dealing with multiple conditions in the same column.

---

## 3. The Literal Equivalent: `SUMX` and `FILTER` (Use with Caution)

Sometimes, a scenario arises where `CALCULATE` won't work because logic needs to be evaluated dynamically row-by-row against a measure, not just a static column.

To achieve the _exact_ row-by-row processing that Tableau's `IF` statement uses, DAX relies on "Iterator" functions (functions ending in "X", like `SUMX`).

**Power BI (DAX):**

```dax
Tech Sales Row-by-Row =
SUMX(
    FILTER(
        Sales,
        RELATED(Products[Category]) = "Tech"
    ),
    Sales[Sales]
)
```

> **How it works:** `FILTER` iterates through the `Sales` table row-by-row, uses `RELATED` to look up the Category from the dimension table, and keeps only the "Tech" rows. `SUMX` then adds up the remaining rows.
>
> **Warning:** Use `CALCULATE` 95% of the time. Iterator functions like `SUMX` consume CPU and can severely slow down a dashboard if used unnecessarily on massive fact tables.
