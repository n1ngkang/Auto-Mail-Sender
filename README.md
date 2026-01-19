# Executive Summary
A high-efficiency email automation engine developed in Google Apps Script (GAS) to streamline the contract termination process for the Hong Kong market. 
This tool automates personalized notification emails, bridging the gap between operational data in Google Sheets and communication templates in Google Docs, significantly reducing manual effort and human error.

# Core Engine Logic
* **Decoupled Template Management**: Unlike traditional hard-coded solutions, this engine retrieves email content from a centralized Google Doc. This architecture allows non-technical team members to update email copy effortlessly without touching a single line of code.
* **Context-Aware Logic**: The system dynamically selects between three distinct templates (`Add-ons`, `TMS + Deposit`, or `Standard TMS`) based on the merchant’s service profile and deposit status.
* **Data-Driven Customization**: Automatically extracts merchant-specific details, AM contact info, and system dates, performing real-time variable replacement to generate bespoke notices.

# Technical Highlights
* **Dynamic Markdown Rendering**: Features a custom parser that converts `**text**` into HTML `<b>` tags, ensuring the generated emails are visually structured and professional.
* **Secure Credential Management**: Utilizes GAS Properties Service to store sensitive Document IDs, ensuring zero exposure of private credentials in the public repository.
