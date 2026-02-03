# Executive Summary
This project optimizes the CSM renewal process by integrating an automated emailing system directly into the CSM Renewal Dashboard. 
By leveraging Google Apps Script, we have created a "one-stop" workflow that enables CSMs to verify data, calculate pricing, and dispatch personalized renewal notices without leaving the dashboard. 
This significantly boosts operational efficiency and minimizes human error in data entry.
# Core Engine Logic
* **One-Stop Dashboard Integration:** Unlike fragmented tools, this engine is built directly atop the source-of-truth dashboard. This architecture enables CSMs to trigger professional notices with a single click without switching platforms.
* **State-Driven UI Toggling:** The system features a conditional rendering engine that automatically switches between "Active" (`✓`) and "Placeholder" (`□`) icons, ensuring the email visually reflects the customer's subscription status in real-time.
* **Contextual Revenue Formatting:** Engineered a dynamic string builder that intelligently appends currency tags (`HK$`) only when revenue-impacting data is detected, preventing unprofessional "$0" displays for inactive add-ons.
# Technical Highlights
* **Logic-Based Exception Handling:** Developed a "Zero-to-Blank" data normalization filter to distinguish between deliberate zero-dollar items and unselected services, granting CSMs granular control over label visibility.
