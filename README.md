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

---

##  Key Features

* **Landmark-First Address Logging:** Tailored for localized address capture.
* **4-Digit PIN Verification:** Secure handoff code required from rider to complete delivery state change.
* **Real-time Short-Polling:** Automated status polling every 5 seconds keeps dispatcher and rider views updated dynamically.
* **Immutable Event Log:** Every status transition generates an audit trail in the `Events` tab.