// --- [TASK-BE-4] CAMELCASE HEADER MAPPER FOR FRONTEND ---
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const ridersSheet = ss.getSheetByName("Riders");
  
  const orders = ordersSheet ? ordersSheet.getDataRange().getValues() : [];
  const riders = ridersSheet ? ridersSheet.getDataRange().getValues() : [];

  function cleanHeader(str) {
    return str.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  let parsedOrders = [];
  if (orders.length > 1) {
    const headers = orders[0].map(cleanHeader);
    parsedOrders = orders.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        if (h) obj[h] = row[i];
      });
      // Map explicit fallback properties for the frontend
      obj.orderId = obj.orderid || obj.id || row[0] || '';
      return obj;
    });
  }
  
  let parsedRiders = [];
  if (riders.length > 1) {
    const headers = riders[0].map(cleanHeader);
    parsedRiders = riders.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        if (h) obj[h] = row[i];
      });
      obj.riderId = obj.riderid || obj.id || row[0] || '';
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
    
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Invalid action." }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleCreateOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ordersSheet = ss.getSheetByName("Orders");
  
  if (!ordersSheet) {
    resetDatabase();
    ordersSheet = ss.getSheetByName("Orders");
  }
  
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
    "",         // Assigned Rider ID
    pin,        // PIN
    "",         // Failure Reason
    new Date(), // Created Timestamp
    ""          // Arrival Timestamp
  ]);
  
  logEvent(orderId, "LOGGED", "Retailer", "Order created");
  
  return { success: true, orderId: orderId, pin: pin };
}

// --- SINGLE COMBINED STATUS HANDLER FOR TASKS BE-1, BE-2, BE-3 ---
function handleUpdateStatus(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const ridersSheet = ss.getSheetByName("Riders");

  if (!ordersSheet) {
    return { success: false, message: 'Orders sheet tab not found.' };
  }

  const data = ordersSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { success: false, message: 'No orders exist in sheet.' };
  }

  // Clean headers for exact matching
  const headers = data[0].map(h => h.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, ''));

  function getColIndex(possibleNames) {
    for (let name of possibleNames) {
      const cleanName = name.toLowerCase().replace(/[^a-z0-9]/g, '');
      const idx = headers.indexOf(cleanName);
      if (idx !== -1) return idx;
    }
    return -1;
  }

  const orderIdCol = getColIndex(["Order ID", "orderid", "id"]);
  const statusCol = getColIndex(["Status"]);
  const pinCol = getColIndex(["PIN", "pin"]);
  const phoneCol = getColIndex(["Customer Phone", "Phone", "custphone"]);
  const riderCol = getColIndex(["Assigned Rider ID", "Rider ID", "riderid"]);
  const distCol = getColIndex(["Estimated Distance (km)", "Estimated Distance", "Distance (km)", "distancekm"]);
  const reasonCol = getColIndex(["Failure Reason", "failurereason"]);
  const arrivalCol = getColIndex(["Arrival Timestamp", "arrivaltimestamp"]);

  if (statusCol === -1) {
    return { success: false, message: 'Sheet column header mismatch for Status.' };
  }

  let rowIndex = -1;
  let orderRowData = null;
  const targetId = String(payload.orderId).trim().toUpperCase();

  // Search by exact Order ID column first
  for (let i = 1; i < data.length; i++) {
    const sheetId = String(data[i][orderIdCol !== -1 ? orderIdCol : 0]).trim().toUpperCase();
    if (sheetId === targetId) {
      rowIndex = i + 1;
      orderRowData = data[i];
      break;
    }
  }

  // Fallback scan: check all columns in case header index was offset
  if (rowIndex === -1) {
    for (let i = 1; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        if (String(data[i][j]).trim().toUpperCase() === targetId) {
          rowIndex = i + 1;
          orderRowData = data[i];
          break;
        }
      }
      if (rowIndex !== -1) break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: 'Order ID not found: ' + payload.orderId };
  }

  const currentStatus = String(orderRowData[statusCol]).trim().toUpperCase();
  const newStatus = String(payload.newStatus).trim().toUpperCase();

  // Workflow State Machine Enforcement
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

  // Rider Range Guard
  if (newStatus === 'ASSIGNED' && payload.riderId) {
    if (ridersSheet) {
      const riderData = ridersSheet.getDataRange().getValues();
      if (riderData.length > 1) {
        const riderHeaders = riderData[0].map(h => h.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, ''));
        const rIdCol = riderHeaders.indexOf('riderid');
        const rRangeCol = riderHeaders.indexOf('maxrangekm');

        let riderMaxRange = 15;
        for (let r = 1; r < riderData.length; r++) {
          const rId = String(riderData[r][rIdCol !== -1 ? rIdCol : 0]).trim().toUpperCase();
          if (rId === String(payload.riderId).trim().toUpperCase()) {
            riderMaxRange = parseFloat(riderData[r][rRangeCol !== -1 ? rRangeCol : 4]) || 15;
            break;
          }
        }

        const estDistance = distCol !== -1 ? (parseFloat(orderRowData[distCol]) || 0) : 0;
        if (estDistance > riderMaxRange) {
          return { success: false, message: `Rider range exceeded! Distance (${estDistance}km) exceeds max range (${riderMaxRange}km).` };
        }
      }
    }

    if (riderCol !== -1) {
      ordersSheet.getRange(rowIndex, riderCol + 1).setValue(payload.riderId);
    }
  }

  // Strict PIN & Manual Override Verification Logic
  if (newStatus === 'DELIVERED') {
    if (payload.isManualOverride) {
      const custPhone = phoneCol !== -1 ? String(orderRowData[phoneCol] || '').trim() : '';
      const expectedPhoneTail = custPhone.slice(-4);
      if (String(payload.overrideCode).trim() !== expectedPhoneTail && !payload.dispatcherApproved) {
        return { success: false, message: 'Manual override failed: Phone tail digits do not match.' };
      }
      if (reasonCol !== -1) {
        ordersSheet.getRange(rowIndex, reasonCol + 1).setValue(`OVERRIDE: ${payload.overrideReason || "Lost PIN"}`);
      }
    } else {
      const storedPin = pinCol !== -1 ? String(orderRowData[pinCol] || '').trim() : '';
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