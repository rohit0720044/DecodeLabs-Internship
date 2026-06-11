-- SQL Data Analytics Project
-- Dataset: ecommerce orders
-- Database table: orders

-- 1. View the first 20 records.
SELECT
    order_id,
    order_date,
    customer_id,
    product,
    quantity,
    unit_price,
    payment_method,
    order_status,
    total_price
FROM orders
LIMIT 20;

-- 2. Find high-value shipped orders using WHERE and ORDER BY.
SELECT
    order_id,
    order_date,
    customer_id,
    product,
    quantity,
    total_price
FROM orders
WHERE order_status = 'Shipped'
  AND total_price >= 2000
ORDER BY total_price DESC
LIMIT 10;

-- 3. Count all orders and calculate total and average sales.
SELECT
    COUNT(*) AS total_orders,
    SUM(total_price) AS total_revenue,
    AVG(total_price) AS average_order_value,
    AVG(quantity) AS average_quantity
FROM orders;

-- 4. Group sales by product.
SELECT
    product,
    COUNT(*) AS order_count,
    SUM(quantity) AS units_sold,
    SUM(total_price) AS revenue,
    AVG(total_price) AS average_order_value
FROM orders
GROUP BY product
ORDER BY revenue DESC;

-- 5. Compare order status totals.
SELECT
    order_status,
    COUNT(*) AS order_count,
    SUM(total_price) AS revenue
FROM orders
GROUP BY order_status
ORDER BY order_count DESC;

-- 6. Compare revenue by referral source.
SELECT
    referral_source,
    COUNT(*) AS orders,
    SUM(total_price) AS revenue,
    AVG(total_price) AS average_order_value
FROM orders
GROUP BY referral_source
ORDER BY revenue DESC;

-- 7. Compare payment methods.
SELECT
    payment_method,
    COUNT(*) AS order_count,
    SUM(total_price) AS revenue,
    AVG(total_price) AS average_order_value
FROM orders
GROUP BY payment_method
ORDER BY revenue DESC;

-- 8. Monthly revenue trend.
SELECT
    strftime('%Y-%m', order_date) AS order_month,
    COUNT(*) AS orders,
    SUM(total_price) AS revenue,
    AVG(total_price) AS average_order_value
FROM orders
GROUP BY strftime('%Y-%m', order_date)
ORDER BY order_month;

-- 9. Products with average order value above 1000.
SELECT
    product,
    COUNT(*) AS order_count,
    AVG(total_price) AS average_order_value
FROM orders
GROUP BY product
HAVING AVG(total_price) > 1000
ORDER BY average_order_value DESC;

-- 10. Orders that used a coupon code.
SELECT
    coupon_code,
    COUNT(*) AS order_count,
    SUM(total_price) AS revenue,
    AVG(total_price) AS average_order_value
FROM orders
WHERE coupon_code IS NOT NULL
  AND coupon_code <> ''
GROUP BY coupon_code
ORDER BY revenue DESC;
