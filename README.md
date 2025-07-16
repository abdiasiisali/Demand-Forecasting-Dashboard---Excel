# Demand-Forecasting-Dashboard-Excel
<img width="1920" height="999" alt="Thumbnail" src="https://github.com/user-attachments/assets/6eb4d729-3c91-473d-b398-4b445e4e9cbf" />


## Overview

This repository contains an advanced demand forecasting dashboard architected in Microsoft Excel to support Sales & Operations Planning (S&OP) processes. The solution is designed for Demand Planners, Sales Analysts, and Management Teams to streamline forecast updates, analyze demand patterns, and generate reliable, data-driven insights to improve decision-making.

The workbook integrates historical sales data, applies multiple forecasting models, and presents the output in an intuitive and interactive dashboard environment.

---

## Key Features

**Multi-Model Forecasting Engine:** Users can dynamically select from a suite of four distinct forecasting models to generate a 12-month demand forecast.
**Interactive Visualisations:** A dynamic line chart renders historical demand (solid blue line) against the forecasted demand (dashed orange line), providing an immediate visual understanding of performance.
**Granular Filtering:** The dashboard includes dropdown filters that allow for deep-dive analysis by individual Item Code, Customer Segment, or any combination thereof.
* **Performance Measurement:** Model performance is quantified using the **Weighted Average Percentage Error (WAPE)**, which measures the average error in proportion to demand volume, giving more weight to high-volume items.

---

## Technical Architecture & Modeling

The solution leverages a hybrid architecture to combine user-friendly interaction with powerful statistical analysis.

**Frontend:** A user interface built entirely in **MS Excel** serves as the front end for data input, filtering, and visualisation. The dashboard is delivered as a macro-enabled workbook.
**Backend Processing:** Advanced time-series analyses for the ARIMA and Exponential Smoothing models are powered by a **Python** backend. This allows for complex calculations that are seamlessly integrated into the Excel environment.

### Forecasting Models Implemented

The dashboard allows users to select from the following models:
1.  **6-Month Moving Average (6M Moving AVG):** A baseline moving average model.
2.  **Excel Built-In:** The native forecasting model available within Excel.
3.  **ARIMA:** An Autoregressive Integrated Moving Average model for sophisticated time-series analysis.
4.  **Exponential Smoothing:** A Holt-Winters Exponential Smoothing model featuring automatic trend and seasonality detection.

---

## User Workflow

The process for generating a forecast is streamlined into three steps:

1. **Data Ingestion:** The user navigates to the **'Orders' sheet** and pastes historical customer orders data, ensuring the date format is DD/MM/YYYY.
2. **Model Selection:** On the main **'Dashboard' sheet**, the user selects a forecast model from the dropdown menu. Python-powered models may take 2-5 seconds to process.
3. **Analysis & Interpretation:** The dashboard automatically updates all visualisations and tables. The user can then analyze the 12-month forecast table and the WAPE metric to assess accuracy.

---

## Getting Started
1. To ensure ease of use, review the user guide included which explains how to operate the dashboard and interpret the results.
2.  Download the macro-enabled workbook from this repository.
3.  Open the file using a compatible version of Microsoft Excel.
4.  When prompted, **ensure that macros are enabled** to allow for full functionality.
