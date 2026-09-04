# PARS (Portfolio Analysis & Review System)

## Project Overview

PARS is a modular, locally hosted web application designed to analyze prospective clients' fixed-income portfolios and compare them against Genter Capital Management's proprietary strategies. By automating data ingestion, enriching holdings with real-time market data, and applying tax-adjusted yield calculations, PARS provides the team with a streamlined tool to generate actionable, side-by-side investment proposals.

## Core Objectives

### 1. Automated Data Ingestion & Entity Resolution

* **Flexible Parsing:** Accept and normalize varying client portfolio formats (Excel/CSV) into a standardized internal schema.
* **Entity Matching:** Utilize Levenshtein distance algorithms (via `RapidFuzz`) to accurately map messy client holdings to the internal fixed-income database.

### 2. Seamless Market Data Integration

* **Batch Processing:** Implement a robust context-managed `blpapi` pipeline to query Bloomberg reference data in bulk, eliminating single-CUSIP bottlenecks.
* **Analytic Enrichment:** Automatically populate critical fixed-income metrics including Yield to Worst (YTW), Yield to Maturity (YTM), Modified/Effective Duration, Convexity, and Credit Ratings (S&P/Moody's).

### 3. Tax-Adjusted Comparative Logic

* **Taxable Equivalent Yield (TEY):** Dynamically compute TEY across the portfolio to ensure fair, level-playing-field comparisons between taxable corporates/treasuries and tax-exempt municipal bonds.
* **Strategy Benchmarking:** Evaluate client allocations against configurable target templates for Genter strategies (e.g., Municipal Quality Intermediate), highlighting discrepancies in sector exposure, maturity distribution, and credit quality.

### 4. Modern, Decoupled Architecture

* **API Layer:** Build a stateless, easily documented backend using FastAPI to handle routing and business logic independently of the UI.
* **Interactive Frontend:** Develop a responsive, single-page application using React/Next.js and Tailwind CSS for drag-and-drop uploads and side-by-side analytical dashboards.
* **Containerized Deployment:** Dockerize the entire application to guarantee flawless local execution across the team's machines without dependency conflicts.

## Technical Stack

* **Backend:** Python, FastAPI
* **Data Engineering:** Pandas, RapidFuzz, Bloomberg BLPAPI
* **Frontend UI:** React, Next.js, Tailwind CSS
* **Infrastructure:** Docker

## Development Milestones

- [ ] **Phase 1: Backend & ETL Pipeline**
  - Define Pydantic validation models for `Bond` and `Portfolio`.
  - Finalize the `BloombergClient` context manager for batch BLPAPI requests.
  - Implement TEY calculation logic and strategy template configurations.
- [ ] **Phase 2: API Integration & Frontend**
  - Expose analysis and aggregation endpoints via FastAPI.
  - Build the React/Next.js interface with Tailwind data tables.
  - Implement report export functionality (PDF generation).

## Data to be included

**For Muni and Taxible:**

- Duration Comparison (Range breakdown in pie chart)
- Maturity Comparison (Range breakdown in pie chart)
- Option Comparison (Break down callable and not)
- Revenue Source/Sector Comparison
- Coupon Rate Comparison
- Credit Rating Comparison
- MUNI: State Breakdown & Taxable Equivalent yield for each
- Top 5 issuer concentration
- Cash Flow & Annual income Projection
- Structure Profile
- Dollar values over percentages (stronger to say you get an extra few thousand a year than a 30 bp improvement)
- Interest rate shock matrix - Show client how each portfolio acts in extreme cases

**Big Feature: AI Insight**

- Include a prompt for internal to carry into AI with basic information about entire portfolio to identify a few top credit risks and demonstrate how our portfolio addresses them (could pertain to revenue source, duration, maturity, call risk, etc.)
- Shifts Internal's workload from manual work/fighting a bad system to higher level research
- On internal to analyze AI response, validate, and clean up as needed
