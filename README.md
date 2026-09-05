 ## Executive Summary & System Brief
REFLEX is a lightweight, offline-resilient delivery control system designed for small Kenyan retailers who currently coordinate deliveries through fragmented channels such as WhatsApp, phone calls and informal records. It creates a structured delivery workflow connecting **retailers, dispatchers, riders and customers**, providing visibility, accountability, customer safety and resilience from order creation through delivery completion or recovery. The system transforms an informal delivery process into a controlled lifecycle:

**Logged → Assigned → Picked Up → Arrived → Delivered**

When a delivery cannot be completed:

**Failed → Rescheduled / Reassigned / Follow-up**

REFLEX is intentionally designed around the realities of small-scale retail operations, including limited infrastructure, variable connectivity, landmark-based addresses and the need for low-cost deployment.

### Key Business Objectives
* **Rapid Handoff Verification:** Secure delivery state transitions via 4-digit PIN verification.
* **Localized Address Handling:** Landmark-first data capture optimized for unnumbered or informal routing contexts.
* **Real-time State Telemetry:** Automated short-polling loops (5-second intervals) to maintain synchronized state between dispatcher and rider interfaces without heavy infrastructure overhead.
* **Customer Safety & Trust:** Provide customers with relevant delivery information, including the assigned driver's name and vehicle or transport details, together with PIN-based handoff verification.
* **Delivery Visibility:** Provide retailers and dispatchers with structured visibility of delivery status and rider assignment.
* **Delivery Recovery:** Explicitly record failed deliveries and support operational recovery through rescheduling and rider reassignment.
* **Operational Resilience:** Support continued operation during temporary connectivity interruptions through local action queuing and synchronization.

* ## The Problem

Small retailers often coordinate deliveries through WhatsApp, phone calls and informal records.

While these tools are accessible, they can create operational gaps:

- Orders can be missed or incorrectly communicated.
- Retailers may have limited visibility of delivery progress.
- Dispatchers may struggle to coordinate multiple riders and orders.
- Customers may not know who is arriving at their door.
- Rider accountability can be unclear.
- Failed deliveries may not be properly recorded.
- Changes to delivery times require additional communication.
- Poor connectivity can interrupt the delivery workflow.
- Businesses have limited structured data for reviewing delivery performance.

---
  
# REFLEX | Logistics & Delivery Control System

**REFLEX** is a lightweight, offline-resilient delivery and logistics management platform engineered for micro-retailers, dispatchers, and field delivery riders.

---

##  System Architecture & Data Flow

1. **Frontend Interface (`index.html`):** Built with HTML5, Tailwind CSS, and vanilla JavaScript. Features role-based views for Retailers, Dispatchers, and Field Riders.
2. **Offline-First Resilience:** Integrated `localStorage` queue intercepts actions during network disruptions, auto-syncing queued payloads back to Google Sheets upon network restoration.
3. **Backend Middleware (`Code.gs`):** Hosted via Google Apps Script as an HTTP Web App endpoint (`doPost`). Enforces finite-state transitions (`LOGGED` → `ASSIGNED` → `PICKED UP` → `DELIVERED`).
4. **Relational Database (`REFLEX_DB`):** Google Sheets database structured across 4 normalized tabs: `Orders`, `Riders`, `Events`, and `Config`.

---

##  Tech Stack & Dependencies

* **Frontend:** HTML5, Tailwind CSS (CDN), JS ES6 (`Fetch API`, `localStorage`, `setInterval` Short-Polling)
* **Backend:** Google Apps Script (`Code.gs`)
* **Database:** Google Sheets (`REFLEX_DB`)
* **Version Control:** Git & GitHub

---

##  Live Demo & Setup Instructions

1. **Database Setup:**
   * Open Google Sheets and create a spreadsheet titled `REFLEX_DB`.
   * Open **Extensions > Apps Script** and deploy `Code.gs` as a Web App (Access: *Anyone*).
2. **Frontend Deployment:**
   * Clone this repository:
     ```bash
     git clone [https://github.com/kwendoEric/reflex-readiness-sprint.git](https://github.com/kwendoEric/reflex-readiness-sprint.git)
     ```
   * Open `index.html` directly in any web browser.
   * Demo Link hosted via vercel: https://reflex-readiness-sprint-psi.vercel.app/

---

##  Key Features

* **Landmark-First Address Logging:** Tailored for localized address capture.
* **4-Digit PIN Verification:** Secure handoff code required from rider to complete delivery state change.
* **Real-time Short-Polling:** Automated status polling every 5 seconds keeps dispatcher and rider views updated dynamically.
* **Immutable Event Log:** Every status transition generates an audit trail in the `Events` tab.
* **Scheduling:** Allows deliveries to be planned around customer and operational availability.
* **Rescheduling:** Allows delivery timing to be changed without losing the original delivery record.
