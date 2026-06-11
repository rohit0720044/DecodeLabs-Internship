# SQL Data Analytics Project

## Project Goal

This project uses SQL queries to extract insights from an ecommerce orders dataset. It demonstrates SQL fundamentals, including `SELECT`, `WHERE`, `ORDER BY`, `GROUP BY`, and basic aggregations such as `COUNT`, `SUM`, and `AVG`.

## Files

- `ecommerce_orders.db` - SQLite database created from the Excel dataset.
- `queries.sql` - SQL queries for analysis.
- `query_results.csv` - Output from the main analysis queries.

## Dataset

The dataset contains 1,200 ecommerce order records.

Main columns:

- `order_id`
- `order_date`
- `customer_id`
- `product`
- `quantity`
- `unit_price`
- `shipping_address`
- `payment_method`
- `order_status`
- `tracking_number`
- `items_in_cart`
- `coupon_code`
- `referral_source`
- `total_price`

## Table Schema

```sql
CREATE TABLE orders (
    order_id TEXT,
    order_date TEXT,
    customer_id TEXT,
    product TEXT,
    quantity INTEGER,
    unit_price REAL,
    shipping_address TEXT,
    payment_method TEXT,
    order_status TEXT,
    tracking_number TEXT,
    items_in_cart INTEGER,
    coupon_code TEXT,
    referral_source TEXT,
    total_price REAL
);
```

## How To Run

Open a terminal in this project folder and run:

```bash
sqlite3 ecommerce_orders.db < queries.sql
```

You can also open `ecommerce_orders.db` in any SQLite browser and run the queries from `queries.sql`.

## Key Insights

- Total orders: 1,200
- Total revenue: 1,264,761.96
- Average order value: 1,053.97
- Average quantity per order: 2.95
- Top product by revenue: Chair with 195,620.11 revenue
- Top referral source by revenue: Instagram with 275,285.45 revenue
- Highest-revenue month: 2024-06 with 68,068.54 revenue
- Most common order status: Cancelled with 250 orders
- Highest-revenue payment method: Credit Card with 263,847.63 revenue

## SQL Skills Demonstrated

### SELECT

Used to retrieve specific columns such as order ID, product, customer ID, status, and total price.

### WHERE

Used to filter records, for example high-value shipped orders and orders with coupon codes.

### ORDER BY

Used to sort results by revenue, order count, average order value, and total price.

### GROUP BY

Used to summarize data by product, order status, payment method, referral source, coupon code, and month.

### Aggregations

Used `COUNT`, `SUM`, and `AVG` to calculate order totals, revenue, units sold, and average order values.
