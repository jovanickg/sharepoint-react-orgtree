# SharePoint Modern Org Chart Web Part

![SPFx](https://img.shields.io/badge/SPFx-1.18-blue.svg) ![React](https://img.shields.io/badge/React-v16-blue.svg) [![Buy Me A Beer](https://img.shields.io/badge/Donate-Buy%20Me%20A%20Beer-orange.svg)](https://www.buymeacoffee.com/jovanickg) [![Donate](https://img.shields.io/badge/Donate-PayPal-blue.svg)](https://www.paypal.com/paypalme/jovanickg)

## Overview
A responsive and highly customizable Organization Chart web part for SharePoint Online. Unlike standard org charts that require Azure AD, this component reads data directly from a **SharePoint List**, giving you full control over the hierarchy, ranking, and employee data without needing complex API permissions.

## Features
* **SharePoint List Driven:** Completely independent of Azure AD; you control the data in a simple list.
* **Department Grouping:** Automatically groups users into department cards.
* **Smart Hierarchy:** Builds the tree structure based on "Superior" references (Department level).
* **Employee vs. Contractor:** Visually distinguishes contractors (pink accent) from internal employees (blue accent).
* **Rank-Based Sorting:** Supports a custom "Job Rank" column to ensure managers appear before interns.
* **Interactive Controls:**
    * Zoom In / Zoom Out / Reset (1:1)
    * **Fit to Screen**
    * **Print Mode** (Optimized CSS for landscape printing)
    * **Pan/Drag:** Click and drag the canvas to navigate large charts.
* **User Details:** Click any user to open a Fluent UI dialog with:
    * High-res profile photo
    * Direct links to Email and Mobile call
    * **"Chat in Teams"** deep link

## Prerequisites: The Data List
You must create a SharePoint List (e.g., named "Employees") with the following columns. You can map these in the web part settings, but these are the defaults:

| Column Name | Type | Description |
| :--- | :--- | :--- |
| **Title** | Single Line of Text | Employee Name |
| **Job Title** | Single Line of Text | e.g., "Senior Developer" |
| **Department** | Single Line of Text | The department this user belongs to. |
| **Superior Department**| Single Line of Text | The **Department Name** of the parent node (e.g., "IT" reports to "Board"). |
| **Contract Type** | Single Line of Text | Used for filtering/styling. Values: "Employee", "Contractor". |
| **Job Rank** | Single Line of Text | (Optional) Sorting code. e.g., "01" for Boss, "99" for Intern. |
| **Email** | Person or Group | The user's system account (used for fetching the Avatar/Photo). |
| **Mobile** | Single Line of Text | Phone number. |

## Installation

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/jovanickg/spfx-modern-org-chart.git](https://github.com/jovanickg/spfx-modern-org-chart.git)
    cd spfx-modern-org-chart
    ```
2.  **Install dependencies:**
    ```bash
    npm install
    ```
3.  **Bundle and Package:**
    ```bash
    gulp bundle --ship
    gulp package-solution --ship
    ```
4.  **Upload:**
    Upload the `.sppkg` file from `sharepoint/solution` to your **App Catalog**.

## Configuration
Add the web part to a SharePoint page and configure the **Property Pane**:

1.  **List Name:** Enter the exact name of your SharePoint List (e.g., `Employees`).
2.  **Column Mappings:** If your list uses different internal names, map them here:
    * *Department Column*
    * *Superior Column*
    * *Job Rank Column* (crucial for ordering branches correctly)
3.  **Contract Type Filter:** Comma-separated list of contract types to include as "Employees" (others are treated as "Contractors").

## Tech Stack
* **Framework:** SharePoint Framework (SPFx)
* **Libraries:**
    * `@pnp/sp` (Data Access)
    * `react-organizational-chart` (Visualization)
    * `@fluentui/react` (UI Components)

## License
MIT

---

### Support
If you find this component useful, consider buying me a beer!

<a href="https://www.buymeacoffee.com/jovanickg" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-orange.png" alt="Buy Me A Beer" style="height: 60px !important;width: 217px !important;" ></a>
