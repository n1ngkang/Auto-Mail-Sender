# Project Overview
This repository contains a collection of automation tools designed to optimize business communication workflows. These scripts handle high-stakes merchant notifications, ranging from contract terminations to service renewals.

# Architectural Strategy
I have implemented two distinct templating methodologies within this suite, chosen based on the frequency of content updates and long-term maintenance needs:
1. Decoupled `Doc-to-Mail` Method (Used in [Termination Mail](./termination-mail/churnmail.js]) )
    * Why: Contract terminations often involve legal wording that requires frequent updates by non-technical stakeholders (Legal/Compliance).
    * Solution: The script retrieves content from a Centralized Google Doc.
    * Benefit: This allows team members to edit the email copy effortlessly in a Document interface without touching the code, ensuring high flexibility and operational safety.

2. Embedded `Script-to-HTML` Method (Used in [Renewal Mail](./renewal-mail/renewalmail.js]) )
    * Why: Renewal notices involve highly structured pricing tables and complex conditional logic (Add-ons, pricing tags, checkmarks) that are closely tied to data columns.
    * Solution: The email structure is hard-coded into HTML templates within the script.
    * Benefit: This provides maximum precision for dynamic rendering and ensures that complex logic-driven UI elements (like the dynamic checkmarks) remain stable and high-performing.

# Setup and Deployment
* Utilizes `.clasp.json` to bridge local development with the Google Apps Script project via the script ID.
* All script logic is maintained in this GitHub repository, using separate directories (`renewal-mail` / `termination-mail`) to prevent global namespace collisions.
* Ensure the manual trigger (menu button) is initialized via the `onOpen()` function to provide CSMs with a one-click mailing interface.
