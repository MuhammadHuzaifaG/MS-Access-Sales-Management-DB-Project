# MS Access Sales Management DB and Querying Project

## Overview

This project implements a relational database structure designed to manage and analyze core business operations, including customer records, product inventory, sales transactions, order details, and sales personnel performance. Using **Microsoft Access** as the database engine, the system centralizes data inputs and provides structured SQL queries for operational reporting and business intelligence.

---

<img src="https://github.com/MuhammadHuzaifaG/MS-Access-Sales-Management-DB-Project/blob/main/Dashboard.PNG" alt="Description of Image" width="1000" height="auto">

## Project Description

### Problem
Organizations often manage transactional data across fragmented spreadsheets, resulting in data redundancy, inconsistent customer records, and difficulty tracking inventory turnover and salesperson performance. Without a relational model, generating comprehensive reports on revenue and stock levels requires manual aggregation.

### Solution
A normalized relational database schema comprising five interconnected tables (`Customers`, `Orders`, `Order Details`, `Products`, and `SalesPersons`) was established. Foreign key constraints link orders to specific customers and sales personnel, while order line items connect transactions to products. Custom SQL queries were developed to extract and aggregate sales performance metrics, customer histories, and product inventory levels directly within Microsoft Access.

### Key Results
- Centralized transaction tracking, eliminating data duplication between customer orders and product inventory.
- Automated aggregation of total revenue per product and total quantity sold using SQL grouping and aggregate functions (`Sum`, `Nz`).
- Improved visibility into sales force activity and order fulfillment metrics.

---

## Database Schema and Data Context

The relational model consists of five primary tables:

1. **Customers**: Stores client contact details and geographic information.
   - *Columns:* `CustomerID`, `FirstName`, `LastName`, `Address`, `City`, `PhoneNumber`
2. **Orders**: Records transactional metadata linking orders to customers and sales representatives.
   - *Columns:* `OrderID`, `CustmerID`, `OrderDate`, `SalesPerson`
3. **Order Details**: Acts as a junction table capturing line-item specifics for each order.
   - *Columns:* `OrderID`, `ProductID`, `Quantity`
4. **Products**: Maintains inventory data, pricing, and stock status.
   - *Columns:* `ProductID`, `ProductName`, `ProductStock`, `UnitsOnOrder`, `UnitPrice`
5. **SalesPersons**: Tracks internal staff details and employment history.
   - *Columns:* `SalesPerson`, `FirstName`, `LastName`, `WorkPhone`, `HireDate`

<img src="https://github.com/MuhammadHuzaifaG/MS-Access-Sales-Management-DB-Project/blob/main/Relationship.PNG" alt="Description of Image" width="1000" height="auto">

---

## Technology Stack

- **Database Management System:** Microsoft Access
- **Query Language:** Structured Query Language (SQL)
- **Data Staging:** Microsoft Excel

---

## Prerequisites

- Microsoft Office Suite (Microsoft Access installed)
- Basic understanding of relational database concepts and SQL syntax

---

## Analytical Results Summary

Below is a summary structure of key operational metrics derived from the database queries:

| Metric Category | Target Entity | Key Indicator | Analytical Purpose |
| :--- | :--- | :--- | :--- |
| Customer Orders | Customers & Orders | Order History & Contact Info | Tracing which customer placed specific orders and when |
| Sales Analysis | Orders & Products | Line-item breakdown | Evaluating product movement across individual orders |
| Product Performance | Products | Total Revenue & Quantity Sold | Identifying top-performing inventory items and revenue generators |

---

## Custom SQL Queries

<img src="https://github.com/MuhammadHuzaifaG/MS-Access-Sales-Management-DB-Project/blob/main/screenshots/Sales%20Analysis%20Query.PNG" alt="Description of Image" width="1000" height="auto">

### 1. Customer's Orders
This query retrieves comprehensive order information combined with customer demographic and contact data using an inner join.

```sql
SELECT 
    Orders.OrderID, 
    Orders.CustmerID, 
    Orders.OrderDate, 
    Orders.SalesPersonID, 
    Customers.FirstName, 
    Customers.LastName, 
    Customers.City, 
    Customers.Address, 
    Customers.PhoneNumber
FROM Customers 
INNER JOIN Orders 
ON Customers.CustomerID = Orders.CustmerID;
```

### 2. Sales Analysis
This query maps line-item order details against products and customer information to provide a granular view of transactions.

```sql
SELECT 
    Orders.OrderID, 
    Orders.OrderDate, 
    Customers.FirstName, 
    Products.ProductName, 
    Products.UnitPrice, 
    [Order Details].Quantity
FROM Customers 
INNER JOIN (Orders 
INNER JOIN (Products 
INNER JOIN [Order Details] 
ON Products.ProductID = [Order Details].ProductID) 
ON Orders.OrderID = [Order Details].OrderID) 
ON Customers.CustomerID = Orders.CustmerID;
```

### 3. Product List
This query calculates total quantities sold and total revenue generated per product, handling null values via the `Nz` function and ordering results by descending revenue.

```sql
SELECT 
    p.ProductID, 
    p.ProductName, 
    p.ProductStock, 
    p.UnitsOnOrder, 
    p.UnitPrice, 
    Nz(Sum(od.Quantity), 0) AS TotalQuantitySold, 
    Nz(Sum(od.Quantity * p.UnitPrice), 0) AS TotalRevenueGenerated
FROM Products AS p 
LEFT JOIN [Order Details] AS od 
ON p.ProductID = od.ProductID
GROUP BY 
    p.ProductID, 
    p.ProductName, 
    p.ProductStock, 
    p.UnitsOnOrder, 
    p.UnitPrice
ORDER BY Nz(Sum(od.Quantity * p.UnitPrice), 0) DESC;
```
