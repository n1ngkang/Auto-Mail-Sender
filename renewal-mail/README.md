# Executive Summary
This project optimizes the CSM renewal process by integrating an automated emailing system directly into the CSM Renewal Dashboard. 
By leveraging Google Apps Script, we have created a "one-stop" workflow that enables CSMs to verify data, calculate pricing, and dispatch personalized renewal notices without leaving the dashboard. 
This significantly boosts operational efficiency and minimizes human error in data entry.
# Core Technical Contributions
## Dynamic Price Tag Logic 
* **Automated Status Visualization:** Implemented a conditional rendering engine that evaluates cell values to dynamically switch between "checked" (`✓`) and "unchecked" (`□`) states, providing clear visual confirmation of subscribed services.
* **Contextual Currency Formatting:** Developed an automated string builder that intelligently appends currency tags ($) only when revenue-impacting data is present, preventing awkward $0 displays for unselected services.
## One-Stop Workflow Integration
* **Cross-Sheet Data Synchronization:** Established a centralized data pipeline that pulls real-time information from the CSM Renewal Dashboard into the automated mailing sheet, eliminating the need for manual copy-pasting and ensuring data integrity.
* **HTML Email Templating:** Developed a robust HTML generation engine within Apps Script that converts spreadsheet rows into professionally styled, mobile-responsive email templates with consistent branding and layout.
## Exception & Edge-Case Management
* **Zero-Value Filtering (Logic Normalization):** Solved the "0 vs. Blank" data ambiguity by creating a specific filter that distinguishes between a deliberate zero and an empty intent, giving CSMs precise control over price label visibility.
# Setup and Deployment
* Utilizes .clasp.json to bridge local development with the Google Apps Script project via the script ID.
* All script logic is maintained in this GitHub repository, using separate directories (`renewal-mail` / `termination-mail`) to prevent global namespace collisions.
* Ensure the manual trigger (menu button) is initialized via the `onOpen()` function to provide CSMs with a one-click mailing interface.
