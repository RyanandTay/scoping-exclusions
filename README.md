# Financial Scoping Exclusions Automation

A VBA-based Excel automation tool designed to streamline the quarterly audit scoping process for internal audit teams. This macro automatically highlights and totals Financial Statement Line Item (FSLI) exclusions across business units and scoping components, eliminating hours of manual formatting work.

## Project Overview

During quarterly audits, internal audit management works with external auditors and the segment VPs to determine which FSLIs should be excluded from audit testing for various Business Units. This exclusion list is constantly revised throughout the quarter and sometimes even after quarter-end based on discussions with compliance professionals.

**The Problem:** Before this automation, team members had to manually update pivot table highlighting and calculate totals each time the exclusion list changed—a time-consuming and error-prone process.

**The Solution:** This VBA macro reads the exclusion list from the "Input" sheet, automatically highlights matching cells in the pivot table on the "Scoping" sheet, and adds calculated totals and percentages. What once took hours of manual work now runs in seconds.

<!-- TODO: Add screenshot of the final formatted pivot table with highlighted exclusions -->

## Features

- **Automated Cell Highlighting**: Matches exclusions from the Input sheet and applies color-coded formatting to the pivot table
- **Flexible Exclusion Logic**: Supports multiple exclusion types:
  - Individual FSLI exclusions for specific Business Units
  - All Balance Sheet FSLIs for a Business Unit
  - All Income Statement FSLIs for a Business Unit
  - All FSLIs for a Business Unit and Scoping Component combination
  - Dynamically adjusts for added or removed FSLIs and Scoping Components
- **Dynamic Totals Calculation**: Automatically calculates and inserts:
  - Total excluded amount per FSLI row
  - Percentage of total excluded
- **Easy Reset**: One-click button to clear all formatting and added columns
- **Handles Large Datasets**: Processes up to 9,000 rows of exclusion efficiently

## Business Context

### What is a Scoping Component?

A **Scoping Component** is the combination of two organizational dimensions:

1. **Business Segment** (one of three high-level organizations):
   - **Consulting**: Professional services and advisory
   - **Facilities**: Property management and operations
   - **Investments**: Asset management and investment operations

2. **Region**: Geographic location (e.g., North America, EMEA, APAC)

**Example:** "Consulting - North America" or "Facilities - EMEA" would each be a unique Scoping Component.

These combinations are important because audit testing scope is often determined at this granular level, requiring management to track exclusions by both segment and region.

## How It Works

### Core Logic

1. **User Input**: Prompts for the row number where Income Statement FSLIs begin (to differentiate from Balance Sheet items)

2. **Exclusion Matching**: Creates formulas in a temporary sheet that use `COUNTIFS()` to check if each pivot table cell matches any exclusion criteria from the Input sheet

3. **Highlighting Application**:
   - First pass: Identifies individual cells that should be excluded from the pivot table or cells where the entire Business Unit is excluded for a Scoping Component (all FSLIs)
   - Second pass: Applies more granular logic for Balance Sheet vs. Income Statement specific exclusions
   - Applies a bluish-gray color (`RGB(220, 220, 240)`) to all matched cells

4. **Totals Calculation**: Loops through each row and sums all highlighted cell values, then:
   - Inserts total excluded amount in a new column
   - Calculates and displays percentage of total excluded

5. **Cleanup**: Deletes the temporary calculation sheet and returns focus to the main Scoping sheet

<!-- TODO: Add screenshot of the Input sheet showing the exclusion list structure -->

## Setup Instructions

### Prerequisites

1. **Enable Macros**:
   - When opening the workbook, click "Enable Content" in the yellow security banner
   - If you don't see the banner, go to File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros (or add this workbook to Trusted Locations)

<!-- TODO: Add screenshot showing the three required sheets in the Excel workbook tabs -->

### Understanding the Buttons

- **Large Green Button**: Runs the main `Scoping_Exclusions` macro
- **Smaller Button**: Runs the `Clean_Scoping` reset macro

<!-- TODO: Add screenshot highlighting the macro buttons location -->

## Usage Steps

### Running the Exclusions Macro

1. **Prepare Your Data**:
   - Ensure the "Original Data" sheet contains your complete financial dataset
   - Update the "Input" sheet with the current list of exclusions
   - Create a pivot table on the "Scoping" sheet from the Original Data

2. **Run the Macro**:
   - Click the **large green button** on the worksheet
   - An input dialog will appear asking: *"Enter the row number where the IS FSLIs starts"*

3. **Provide Income Statement Row**:
   - Look at your pivot table and identify which row the Income Statement data begins
   - Enter a number greater than 10 (validation will prompt if invalid)
   - Click OK

4. **Wait for Processing**:
   - The macro will process (screen updating is disabled for speed)
   - A message box will confirm when complete: *"The Macro has finished running"*

5. **Review Results**:
   - Excluded cells will be highlighted in bluish-gray
   - Two new columns appear after the pivot table:
     - "Total Excluded by FSLI"
     - "Percentage of total Excluded"

<!-- TODO: Add before/after screenshots showing pivot table transformation -->

**Future Enhancements** could include:
- Error handling (`On Error GoTo ErrorHandler`)
- Function decomposition for better modularity
- Unit testing framework
- Configuration sheet for user-definable constants
- Progress indicator for large datasets

## License

This project was created for internal business use.

## Author

**Ryan Neilson**
Developed during tenure as an accountant, demonstrating problem-solving and automation skills applied to real business workflows.

---

*This automation reduced quarterly scoping update time from hours to seconds, directly improving team efficiency and reducing manual errors in the audit process.*
