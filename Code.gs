function getColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    map[header.toString().trim()] = index + 1; // 1-based index for Apps Script API
  });
  return map;
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const ridersSheet = ss.getSheetByName("Riders");
  
  const orders = ordersSheet ? ordersSheet.getDataRange().getValues() : [];
  const riders = ridersSheet ? ridersSheet.getDataRange().getValues() : [];
  
  let parsedOrders = [];
  if (orders.length > 1) {
    const headers = orders[0];
    parsedOrders = orders.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
  }
  
  let parsedRiders = [];
  if (riders.length > 1) {
    const headers = riders[0];
    parsedRiders = riders.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
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
    else if (payload.action === 'UPDATE_STATUS' || payload.action === 'VERIFY_PIN_DELIVER') {
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
    "LOGGED",   // Status
    "",         // Assigned Rider ID
    pin,        // PIN
    "",         // Failure Reason
    new Date(), // Created Timestamp
    ""          // Arrival Timestamp
  ]);
  
  logEvent(orderId, "LOGGED", "Retailer", "Order created");
  
  return { success: true, orderId: orderId, pin: pin };
}

function handleUpdateStatus(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const data = ordersSheet.getDataRange().getValues();
  const colMap = getColumnMap(ordersSheet);
  
  const orderId = payload.orderId;
  const newStatus = payload.newStatus || payload.status;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      const rowIndex = i + 1; // 1-based sheet row index
      
      // Dynamically update fields based on column headers
      if (newStatus && colMap["Status"]) {
        ordersSheet.getRange(rowIndex, colMap["Status"]).setValue(newStatus);
      }
      
      if (payload.riderId && colMap["Assigned Rider ID"]) {
        ordersSheet.getRange(rowIndex, colMap["Assigned Rider ID"]).setValue(payload.riderId);
      }
      
      if (payload.failureReason && colMap["Failure Reason"]) {
        ordersSheet.getRange(rowIndex, colMap["Failure Reason"]).setValue(payload.failureReason);
      }
      
      if (newStatus === "ARRIVED" && colMap["Arrival Timestamp"]) {
        ordersSheet.getRange(rowIndex, colMap["Arrival Timestamp"]).setValue(new Date());
      }
      
      logEvent(orderId, newStatus || "UPDATED", payload.actorRole || "SYSTEM", payload.failureReason || "");
      
      return { success: true, message: `Order ${orderId} updated successfully` };
    }
  }
  
  return { success: false, message: "Order ID not found" };
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
  
  // Setup Orders Sheet
  let ordersSheet = ss.getSheetByName("Orders");
  if (!ordersSheet) ordersSheet = ss.insertSheet("Orders");
  ordersSheet.clear();
  ordersSheet.appendRow([
    "Order ID", "Shop Name", "Shop Type", "Verified Retailer", "Pickup Location",
    "Item", "Item Model", "Item Qty", "Customer Name", "Phone", "Area", "Landmark",
    "Distance (km)", "Out of Zone", "Delivery Type", "Is Urgent", "Delivery Date",
    "Time Slot", "Status", "Assigned Rider ID", "PIN", "Failure Reason",
    "Created Timestamp", "Arrival Timestamp"
  ]);
  
  // Setup Riders Sheet
  let ridersSheet = ss.getSheetByName("Riders");
  if (!ridersSheet) ridersSheet = ss.insertSheet("Riders");
  ridersSheet.clear();
  ridersSheet.appendRow(["riderId", "name", "phone", "vehicle", "maxRangeKm", "operatingZone", "status"]);
  
  ridersSheet.appendRow(["RDR-001", "John Kimani", "0711111111", "Motorbike", 15, "Nairobi CBD / Central", "ACTIVE"]);
  ridersSheet.appendRow(["RDR-002", "Peter Otieno", "0722222222", "Bicycle", 8, "Westlands / Kilimani", "ACTIVE"]);
  ridersSheet.appendRow(["RDR-003", "Mercy Wanjiku", "0733333333", "Van / Pickup", 50, "All Nairobi Zones", "ACTIVE"]);
}