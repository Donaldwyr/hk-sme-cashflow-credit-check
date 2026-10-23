# Project Report: Hong Kong SME Cash Flow & Credit Risk Screening Tool

## 1. Introduction

In Hong Kong’s commercial banking industry, SME trade and working capital financing constitute one of the most core and high-frequency business lines for Relationship Managers (RMs). Before formal field visits, detailed due diligence, and official credit proposal writing, frontline banking teams require a fast, standardized, and reliable preliminary screening mechanism to judge whether a corporate borrower’s financial status is acceptable for further credit review.

Traditional manual financial ratio calculation suffers from obvious drawbacks: repetitive manual Excel computation, inconsistent accounting caliber, human calculation errors, and inefficient trend comparison across fiscal years. This project builds a fully automated Python-based SME credit screening system tailored to **Hong Kong commercial banking risk assessment standards**. The tool automatically reads financial statements, calculates core credit metrics, generates visualized financial trend charts, outputs risk ratings, and produces structured preliminary credit reports.

The tool is designed specifically for frontline RM preliminary risk judgment scenarios. It focuses on evaluating SME short-term liquidity, long-term leverage level, interest-bearing repayment capacity, and real operating cash flow quality, simulating the standard pre-screening workflow adopted by Hong Kong commercial banks for SME working capital loan applications.

## 2. Methodology & Credit Indicator Design

The system strictly selects five core financial indicators that are most valued in Hong Kong SME credit reviews, covering liquidity risk, leverage risk, profitability repayment capacity, and cash flow robustness. All threshold standards fully align with HSBC Hong Kong commercial banking SME credit risk management rules.

### 2.1 Core Financial Metrics and Economic Meaning

1. **Current Ratio**
Formula: Current Assets / Current Liabilities
It reflects the short-term solvency capability and overall liquidity level of an enterprise. Hong Kong trade SMEs usually hold large inventory and receivables, making this ratio the primary judgment indicator for short-term capital chain stability.
2. **Quick Ratio**
Formula: (Current Assets − Inventory) / Current Liabilities
Inventory of Hong Kong trading SMEs has strong industry-specific uncertainty, including slow turnover, price fluctuation, and depreciation risk. The quick ratio excludes inventory assets to evaluate the enterprise’s real immediate solvency without relying on inventory liquidation.
3. **Asset-Liability Ratio**
Formula: Total Liabilities / Total Assets
It measures the long-term leverage level and overall financial risk of the enterprise. Banking institutions strictly control the upper limit of SME leverage to avoid systemic default risks caused by excessive debt stacking.
4. **Interest Coverage Ratio**
Formula: EBIT / Interest Expense
This indicator verifies whether the enterprise’s core operating profit can fully cover financing costs. It is the core criterion for judging the sustainability of corporate debt repayment and avoiding interest default risks.
5. **Adjusted Operating Cash Flow**
Calculated through simplified indirect method adjustment: starting from book profit, adding back non-cash depreciation and interest expenses, then adjusting working capital changes including accounts receivable, inventory, and accounts payable. It eliminates the illusion of book profit and reflects the real cash inflow capacity of SME core business.

### 2.2 Banking Risk Rating Rules

Three-tier rating standards fully compliant with Hong Kong commercial banking practices:

- **Safe**: Fully meets bank risk threshold requirements
- **Watch**: Floating between warning line and safety line, requiring further due diligence and risk mitigation measures
- **Alert**: Below risk warning line, with significant credit risk exposure

Comprehensive rating logic follows strict banking risk control principles:

- More than two Alert indicators → Overall Alert rating
- No Alert indicators but more than two Watch indicators → Overall Watch rating
- Any single Alert indicator → Minimum Watch rating to avoid underrating risks
- All indicators meet safety standards → Overall Safe rating

## 3. Data Structure & Modeling Assumptions

### 3.1 Data Source Structure

The project adopts standardized Hong Kong SME financial statement templates, containing two core sheets: Balance Sheet (BS) and Income Statement (IS), covering three consecutive fiscal years (FY2021–FY2023). All financial data units are thousand Hong Kong dollars, consistent with the financial statement presentation habits of Hong Kong small and medium-sized trading enterprises.

### 3.2 Key Modeling Assumptions

In line with the actual situation of incomplete financial disclosure of Hong Kong SMEs:

1. No significant non-operating income and non-operating expenses, EBIT is approximately equal to operating profit after adding back interest expenses
2. Exclude complex non-operating working capital adjustment items, only retain the three most core adjustment items: receivables, inventory, payables
3. Do not involve fixed asset capital expenditure, equity financing, and investment income items, focusing purely on core business cash flow performance
4. All financial data maintain strict accounting double-entry balance to ensure the authenticity and rigor of financial logic

## 4. Testing Result & Practical Business Interpretation

### 4.1 Baseline Test Result (FY2023)

The verified enterprise presents typical **Watch-level SME customer characteristics in Hong Kong commercial banking**:

- Slightly insufficient short-term liquidity reserve
- Relatively high leverage ratio for trading enterprises
- Profit repayment capacity close to the safety threshold
- Stable and positive core operating cash flow

### 4.2 Banking Business Decision Interpretation

In actual HSBC commercial banking business scenarios, Watch-rated SMEs will not be directly rejected for credit applications. Instead, RMs will adopt targeted risk mitigation measures:

- Require additional personal unlimited guarantee or corporate guarantee
- Appropriately reduce comprehensive credit line scale
- Shorten loan tenure and adopt revolving short-term loan structure
- Strengthen post-loan fund usage supervision and monthly financial reconciliation

This fully conforms to the differentiated risk pricing and risk control strategies of Hong Kong banks for SME customers.

## 5. System Limitations & Risk Boundaries

This tool is positioned as a **preliminary quantitative screening auxiliary system**, not a substitute for manual professional credit review. The core limitations are clearly defined in line with real banking business logic:

1. **Pure quantitative evaluation limitation**: The system only relies on financial statement data and cannot incorporate qualitative indicators including management background, industry cycle, customer concentration, market reputation, and corporate governance quality.
2. **No asset collateral valuation module**: The credit model does not include collateral value, mortgage rate, and asset realizability evaluation, which are core components of Hong Kong SME credit approval.
3. **Simplified cash flow calculation**: The indirect method adopted is a streamlined version without full financial statement item adjustment, which cannot replace formal audited cash flow statements.
4. **Unified industry threshold limitation**: The current risk threshold adopts universal standards for Hong Kong trading SMEs, without differentiated threshold calibration for technology, manufacturing, and service industry enterprises.
5. **Inability to identify financial fraud**: The system calculates based on input financial data and cannot identify false financial statements or manipulated profit data.

## 6. Future Optimization & Iteration Direction

Combined with the actual pain points of commercial banking credit business, the tool has clear scalable iteration space:

1. **Industry-customized risk threshold system**: Set independent safety and warning thresholds for trading, manufacturing, and technology SMEs to improve rating accuracy across industries.
2. **Introduction of classic credit measurement models**: Integrate Z-Score default prediction model to form a dual evaluation system of financial ratio + default probability.
3. **Automatic real financial data crawling**: Connect Hong Kong corporate public financial data interfaces to realize automatic data acquisition without manual Excel entry.
4. **Qualitative scoring card module**: Add weighted scoring dimensions including customer industry attributes, operating years, and credit record to build a comprehensive quantitative + qualitative credit evaluation system.
5. **Stress test module upgrade**: Realize one-click batch scenario simulation including revenue decline, inventory depreciation, interest rate hike, and receivable overdue to test enterprise financial resilience under macro shocks.

## 7. Conclusion

This project successfully builds a lightweight, banking-standard SME credit risk screening tool oriented to Hong Kong commercial banking frontline business scenarios. It standardizes the manual financial calculation and risk rating process in traditional SME credit pre-review, effectively reducing manual calculation errors and improving the efficiency of preliminary customer screening for commercial banking RMs.

All financial formulas strictly follow accounting principles, and risk thresholds fully align with HSBC Hong Kong SME credit risk management specifications. The tool clearly defines application boundaries, takes automated quantitative analysis as the core auxiliary capability, and reserves professional business judgment rights for bankers, which is highly consistent with the actual working logic of Hong Kong commercial banking credit business.
