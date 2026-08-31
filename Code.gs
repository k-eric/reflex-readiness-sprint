// --- [TASK-BE-4] CAMELCASE HEADER MAPPER FOR FRONTEND ---
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const ridersSheet = ss.getSheetByName("Riders");
  
  const orders = ordersSheet ? ordersSheet.getDataRange().getValues() : [];
  const riders = ridersSheet ? ridersSheet.getDataRange().getValues() : [];
  
  // Convert header string to camelCase (e.g. "Order ID" -> "orderId")
  function toCamelCase(str) {
    return str.toString().trim()
      .replace(/(?:^\w|[A-Z]|\b\w)/g, (letter, index) => 
        index === 0 ? letter.toLowerCase() : letter.toUpperCase()
      ).replace(/\s+/g, '').replace(/[^a-zA-Z0-9]/g, '');
  }

  let parsedOrders = [];
  if (orders.length > 1) {
    const headers = orders[0];
    parsedOrders = orders.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => { 
        if (h) obj[toCamelCase(h)] = row[i]; 
      });
      return obj;
    });
  }
  
  let parsedRiders = [];
  if (riders.length > 1) {
    const headers = riders[0];
    parsedRiders = riders.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => { 
        if (h) obj[toCamelCase(h)] = row[i]; 
      });
      return obj;
    });
  }

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    orders: parsedOrders,
    riders: parsedRiders
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    
    if (payload.action === 'CREATE_ORDER') {
      return ContentService.createTextOutput(JSON.stringify(handleCreateOrder(payload)))
        .setMimeType(ContentService.MimeType.JSON);
    } 
    else if (payload.action === 'UPDATE_STATUS') {
      return ContentService.createTextOutput(JSON.stringify(handleUpdateStatus(payload)))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleCreateOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  
  const orderId = "ORD-" + Math.floor(1000 + Math.random() * 9000);
  const pin = Math.floor(1000 + Math.random() * 9000).toString();
  
  ordersSheet.appendRow([
    orderId,
    payload.shopName || "",
    payload.shopType || "",
    payload.isVerifiedRetailer ? "YES" : "NO",
    payload.pickupLocation || "",
    payload.item || "",
    payload.itemModel || "",
    payload.itemQty || 1,
    payload.customerName || "",
    payload.phone || "",
    payload.area || "",
    payload.landmark || "",
    payload.estimatedDistance || 0,
    payload.isOutOfZone ? "YES" : "NO",
    payload.deliveryType || "Immediate",
    payload.isUrgent ? "YES" : "NO",
    payload.deliveryDate || "",
    payload.timeSlot || "",
    "LOGGED",   // status
    "",         // riderId
    pin,        // pin
    "",         // failureReason
    new Date(), // createdTimestamp
    ""          // arrivalTimestamp
  ]);
  
  logEvent(orderId, "LOGGED", "Retailer", "Order created");
  
  return { success: true, orderId: orderId, pin: pin };
}

// --- SINGLE COMBINED STATUS HANDLER FOR TASKS BE-1, BE-2, BE-3 ---
function handleUpdateStatus(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const ridersSheet = ss.getSheetByName("Riders");

  const data = ordersSheet.getDataRange().getValues();
  const headers = data[0];

  function getColIndex(name) {
    return headers.findIndex(h => h.toString().trim() === name);
  }

  const orderIdCol = getColIndex("Order ID");
  const statusCol = getColIndex("Status");
  const pinCol = getColIndex("PIN");
  const phoneCol = getColIndex("Customer Phone");
  const riderCol = getColIndex("Assigned Rider ID");
  const distCol = getColIndex("Estimated Distance (km)");
  const reasonCol = getColIndex("Failure Reason");
  const arrivalCol = getColIndex("Arrival Timestamp");

  if (orderIdCol === -1 || statusCol === -1) {
    return { success: false, message: 'Sheet column header mismatch.' };
  }

  let rowIndex = -1;
  let orderRowData = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][orderIdCol]).trim() === String(payload.orderId).trim()) {
      rowIndex = i + 1;
      orderRowData = data[i];
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: 'Order ID not found: ' + payload.orderId };
  }

  const currentStatus = String(orderRowData[statusCol]).trim().toUpperCase();
  const newStatus = String(payload.newStatus).trim().toUpperCase();

  // [TASK-BE-2] Workflow State Machine Enforcement
  const validTransitions = {
    'LOGGED': ['ASSIGNED', 'ESCALATED', 'FAILED'],
    'ASSIGNED': ['PICKED UP', 'ESCALATED', 'FAILED', 'LOGGED'],
    'PICKED UP': ['ARRIVED', 'ESCALATED', 'FAILED'],
    'ARRIVED': ['DELIVERED', 'ESCALATED', 'FAILED'],
    'ESCALATED': ['ASSIGNED', 'LOGGED', 'FAILED'],
    'FAILED': ['ASSIGNED', 'LOGGED']
  };

  if (validTransitions[currentStatus] && !validTransitions[currentStatus].includes(newStatus)) {
    return { success: false, message: `Invalid transition: Cannot move order from ${currentStatus} to ${newStatus}.` };
  }

  // [TASK-BE-3] Rider Range Guard
  if (newStatus === 'ASSIGNED' && payload.riderId) {
    const riderData = ridersSheet.getDataRange().getValues();
    const riderHeaders = riderData[0];
    const rIdCol = riderHeaders.findIndex(h => h.toString().trim() === 'Rider ID');
    const rRangeCol = riderHeaders.findIndex(h => h.toString().trim() === 'Max Range (km)');

    let riderMaxRange = 15;
    for (let r = 1; r < riderData.length; r++) {
      if (String(riderData[r][rIdCol]).trim() === String(payload.riderId).trim()) {
        riderMaxRange = parseFloat(riderData[r][rRangeCol]) || 15;
        break;
      }
    }

    const estDistance = parseFloat(orderRowData[distCol]) || 0;
    if (estDistance > riderMaxRange) {
      return { success: false, message: `Rider range exceeded! Distance (${estDistance}km) exceeds max range (${riderMaxRange}km).` };
    }

    if (riderCol !== -1) {
      ordersSheet.getRange(rowIndex, riderCol + 1).setValue(payload.riderId);
    }
  }

  // [TASK-BE-1] Strict PIN & Manual Override Verification Logic
  if (newStatus === 'DELIVERED') {
    if (payload.isManualOverride) {
      const custPhone = String(orderRowData[phoneCol] || '').trim();
      const expectedPhoneTail = custPhone.slice(-4);
      if (String(payload.overrideCode).trim() !== expectedPhoneTail && !payload.dispatcherApproved) {
        return { success: false, message: 'Manual override failed: Phone tail digits do not match.' };
      }
      if (reasonCol !== -1) {
        ordersSheet.getRange(rowIndex, reasonCol + 1).setValue(`OVERRIDE: ${payload.overrideReason || "Lost PIN"}`);
      }
    } else {
      const storedPin = String(orderRowData[pinCol] || '').trim();
      const submittedPin = String(payload.pin || '').trim();

      if (!submittedPin || submittedPin !== storedPin) {
        return { success: false, message: 'Invalid verification PIN. Delivery cannot be completed.' };
      }
    }
  }

  if (payload.failureReason && reasonCol !== -1) {
    ordersSheet.getRange(rowIndex, reasonCol + 1).setValue(payload.failureReason);
  }

  if (arrivalCol !== -1 && newStatus === "ARRIVED") {
    ordersSheet.getRange(rowIndex, arrivalCol + 1).setValue(new Date());
  }

  ordersSheet.getRange(rowIndex, statusCol + 1).setValue(newStatus);
  logEvent(payload.orderId, newStatus, payload.actorRole || "SYSTEM", payload.failureReason || payload.overrideReason || "");

  return { success: true, message: `Order ${payload.orderId} updated to ${newStatus}` };
}

function logEvent(orderId, status, actor, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let eventsSheet = ss.getSheetByName("Events");
  if (!eventsSheet) {
    eventsSheet = ss.insertSheet("Events");
    eventsSheet.appendRow(["Timestamp", "Order ID", "Status", "Actor", "Details"]);
  }
  eventsSheet.appendRow([new Date(), orderId, status, actor, details]);
}

function resetDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let ordersSheet = ss.getSheetByName("Orders");
  if (!ordersSheet) ordersSheet = ss.insertSheet("Orders");
  ordersSheet.clear();
  ordersSheet.appendRow([
    "Order ID", "Shop Name", "Shop Type", "Is Verified Retailer", "Pickup Location",
    "Item", "Item Model", "Item Qty", "Customer Name", "Customer Phone", "Area", "Landmark",
    "Estimated Distance (km)", "Is Out Of Zone", "Delivery Type", "Is Urgent", "Delivery Date",
    "Time Slot", "Status", "Assigned Rider ID", "PIN", "Failure Reason",
    "Created Timestamp", "Arrival Timestamp"
  ]);
  
  let ridersSheet = ss.getSheetByName("Riders");
  if (!ridersSheet) ridersSheet = ss.insertSheet("Riders");
  ridersSheet.clear();
  ridersSheet.appendRow(["Rider ID", "Name", "Phone", "Vehicle", "Max Range (km)", "Operating Zone", "Status"]);
  
  ridersSheet.appendRow(["RDR-001", "John Kimani", "0711111111", "Motorbike", 15, "Nairobi CBD / Central", "ACTIVE"]);
  ridersSheet.appendRow(["RDR-002", "Peter Otieno", "0722222222", "Bicycle", 8, "Westlands / Kilimani", "ACTIVE"]);
  ridersSheet.appendRow(["RDR-003", "Mercy Wanjiku", "0733333333", "Van / Pickup", 50, "All Nairobi Zones", "ACTIVE"]);
}