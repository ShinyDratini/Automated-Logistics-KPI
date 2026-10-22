
CREATE DATABASE IF NOT EXISTS logistics_kpi;

USE logistics_kpi;

CREATE TABLE IF NOT EXISTS orders (
    order_id INT PRIMARY KEY,
    order_date DATETIME,
    shipping_date DATETIME,
    delivery_status VARCHAR(50),
    late_delivery_risk TINYINT,
    actual_shipping_days INT,
    scheduled_shipping_days INT,
    shipping_mode VARCHAR(50),
    market VARCHAR(50),
    order_region VARCHAR(100),
    order_country VARCHAR(100),
    order_state VARCHAR(100),
    order_city VARCHAR(100),
    order_status VARCHAR(50)
);

CREATE TABLE IF NOT EXISTS order_items (
    order_item_id INT PRIMARY KEY,
    order_id INT NOT NULL,
    quantity INT,
    sales DECIMAL(12,2),

    FOREIGN KEY (order_id)
        REFERENCES orders(order_id)
);

CREATE INDEX idx_orders_delivery
ON orders(delivery_status);

CREATE INDEX idx_orders_region
ON orders(order_region);
