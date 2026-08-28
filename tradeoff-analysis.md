# REFLEX — One-Page Trade-Off Log

**Project:** Reflex: The Readiness Sprint  
**Author:** Louisa M. Mkhabela (Product Lead / Partner B)  
**Purpose:** This log records the main compromises made while building Reflex within the sprint. The priority was to prove the delivery workflow clearly: a retailer creates a delivery, a dispatcher assigns it, a rider updates it, and the team can see what happened. Simple tools were chosen to support a clear build and explanation within the available time. These choices suit the prototype, but they are not presented as the final production architecture.

---

## 1. Google Sheets as the Datastore

* **Decision:** Google Sheets, named `REFLEX_DB`, was chosen as the datastore for the MVP.
* **What We Gave Up:** Google Sheets has limited support for transactions, concurrent updates, complex queries, and production-level security. It may become unreliable as the number of users and deliveries grows.
* **Why This Was Acceptable:** It was quick to set up, easy to inspect, and worked directly with Google Apps Script. This kept the focus on proving the delivery workflow rather than spending most of the sprint setting up database infrastructure.
* **What We Would Improve Next:** Move to a managed relational database with transactions, constraints, indexes, backups, monitoring, and stronger access control. The existing separation between `Orders`, `Riders`, `Events`, and `Config` would make that migration easier.

---

## 2. Five-Second Polling Instead of Push Updates

* **Decision:** Short polling at approximately five-second intervals was chosen to refresh delivery statuses.
* **What We Gave Up:** Updates are near real-time rather than instant. The system also makes repeated requests, including when no new information is available.
* **Why This Was Acceptable:** The prototype does not need sub-second updates to prove the delivery process. Polling is simple, predictable, and compatible with the Google Apps Script setup.
* **What We Would Improve Next:** Measure update delay and request volume during a pilot. If faster updates or more active users required a different approach, event-driven updates, Server-Sent Events, or WebSockets would be considered.

---

## 3. A Small Browser Queue Using `localStorage`

* **Decision:** A small browser-based queue was chosen to hold rider actions when the connection is temporarily unavailable.
* **What We Gave Up:** Browser storage is limited and less reliable than a proper offline database. Queued information could be lost if browser data is cleared, and conflict handling is basic.
* **Why This Was Acceptable:** The MVP only needs to demonstrate that a rider action can remain pending during a connection problem and synchronize later. Building a complete offline mobile application would have added significant complexity without helping prove the main workflow.
* **What We Would Improve Next:** Use IndexedDB for more reliable local storage. Add retry rules, duplicate protection, and clear handling for conflicting or out-of-order updates.

---

## 4. Simplified Authentication for the Demo

* **Decision:** Simplified authentication and quick role switching were used for the competition demonstration.
* **What We Gave Up:** This does not provide the identity verification, session security, or access control expected in a production system. It would not be enough to protect customer or delivery information at scale.
* **Why This Was Acceptable:** The demo needs to move quickly between the retailer, dispatcher, and rider views. Simplified access keeps the demonstration reliable and keeps the focus on the delivery workflow.
* **What We Would Improve Next:** Add secure login, role-based permissions, session controls, device security, and access logs. Each user would only view and perform actions relevant to their role.

---

## 5. Limited Protection Against Simultaneous Assignments

* **Decision:** A simple assignment process backed by Google Sheets was used for the MVP.
* **What We Gave Up:** Two dispatchers could potentially try to assign the same delivery or rider at almost the same time. The prototype does not provide the full transactional protection expected in a production system.
* **Why This Was Acceptable:** The prototype uses a controlled test dataset and is designed to demonstrate the assignment process. This is identified as a real weakness rather than presented as solved.
* **What We Would Improve Next:** Add transactional assignment, conflict checks, locking or compare-and-set logic, and automated tests for simultaneous actions.

---

## Overall Principle

The simplest approach was chosen to demonstrate the operating model clearly within the sprint. Each decision helped the team move faster, but each one also has a known cost and a clear improvement path.

> **Closing Statement:** Prototype scope is not production readiness. The goal was to prove the workflow, identify its weak points, and show how Reflex could be strengthened in the next phase.
