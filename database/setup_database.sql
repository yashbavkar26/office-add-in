-- ===============================================================
-- Excel to MySQL Auto-Sync Database Setup Script
-- ===============================================================
-- You can run this script directly in MySQL Workbench, phpMyAdmin, or MySQL CLI.
-- Note: The Sync Bridge Server also creates this automatically if needed!

CREATE DATABASE IF NOT EXISTS `excel_sync_db` 
CHARACTER SET utf8mb4 
COLLATE utf8mb4_unicode_ci;

USE `excel_sync_db`;

-- Drop table if you wish to reset
-- DROP TABLE IF EXISTS `orders`;

CREATE TABLE IF NOT EXISTS `orders` (
  `_excel_row_num` INT NOT NULL UNIQUE COMMENT 'Tracks the 1-based row index in Excel to enable in-place row updating',
  `order_id` VARCHAR(50) NULL,
  `customer_name` VARCHAR(255) NULL,
  `product` VARCHAR(255) NULL,
  `category` VARCHAR(100) NULL,
  `quantity` INT NULL,
  `unit_price` DECIMAL(10, 2) NULL,
  `total_amount` DECIMAL(12, 2) NULL,
  `order_status` VARCHAR(50) NULL,
  `order_date` VARCHAR(50) NULL,
  `_created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  `_updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`_excel_row_num`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Sample initial dummy data matching the Excel starter template
INSERT INTO `orders` 
  (`_excel_row_num`, `order_id`, `customer_name`, `product`, `category`, `quantity`, `unit_price`, `total_amount`, `order_status`, `order_date`)
VALUES
  (2, 'ORD-1001', 'Alice Johnson', 'Wireless Mechanical Keyboard', 'Electronics', 2, 89.99, 179.98, 'Delivered', '2026-09-15'),
  (3, 'ORD-1002', 'Brian Smith', 'Ergonomic Office Chair', 'Furniture', 1, 249.50, 249.50, 'Shipped', '2026-09-18'),
  (4, 'ORD-1003', 'Carla Gomez', 'USB-C Multi-Port Hub', 'Accessories', 3, 34.00, 102.00, 'Processing', '2026-09-22'),
  (5, 'ORD-1004', 'David Lee', 'Ultra-Wide Monitor 34-inch', 'Electronics', 1, 499.99, 499.99, 'Delivered', '2026-09-25'),
  (6, 'ORD-1005', 'Emma Watson', 'Noise Cancelling Headphones', 'Audio', 2, 129.00, 258.00, 'Pending', '2026-09-28')
ON DUPLICATE KEY UPDATE
  `customer_name` = VALUES(`customer_name`),
  `product` = VALUES(`product`),
  `quantity` = VALUES(`quantity`),
  `unit_price` = VALUES(`unit_price`),
  `total_amount` = VALUES(`total_amount`),
  `order_status` = VALUES(`order_status`),
  `_updated_at` = CURRENT_TIMESTAMP;
