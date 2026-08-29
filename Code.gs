/**
 * REFLEX BACKEND LOGIC - FIXED & ALIGNED SCHEMAS
 */

// Helper to append a new event to the audit trail
function logEvent(orderId, actor, action, oldStatus, newStatus, device = 'Web', syncStatus = 'Synced') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = ss.getSheetByName('Events');
  const eventId = 'EV-' + new Date().getTime();
  const timestamp = new Date();
  
  eventsSheet.appendRow([eventId, orderId, timestamp, actor, action, oldStatus, newStatus, device, syncStatus]);
}

// Simple state transition validator
function updateOrderStatus(orderId, newStatus, actorRole) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  const data = ordersSheet.getDataRange().getValues();
  
  const validTransitions = {
    'LOGGED': 'ASSIGNED',
    'ASSIGNED': 'PICKED UP',
    'PICKED UP': 'DELIVERED'
  };
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      const currentStatus = data[i][6]; // Column 7: Status (Index 6)
      
      // Verify valid sequence
      if (validTransitions[currentStatus] === newStatus) {
        ordersSheet.getRange(i + 1, 7).setValue(newStatus);
        logEvent(orderId, actorRole, 'STATUS_UPDATE', currentStatus, newStatus);
        return { success: true, message: `Status updated from ${currentStatus} to ${newStatus}` };
      } else {
        return { success: false, message: `Invalid transition from ${currentStatus} to ${newStatus}` };
      }
    }
  }
  return { success: false, message: 'Order ID not found' };
}

// Expose GET endpoint for fetching orders/riders
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orders = ss.getSheetByName('Orders').getDataRange().getValues();
  return ContentService.createTextOutput(JSON.stringify({ success: true, orders: orders }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Expose POST endpoint for creating/updating orders
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (data.action === 'CREATE_ORDER') {
      const ordersSheet = ss.getSheetByName('Orders');
      const orderId = 'RX-' + Math.floor(1000 + Math.random() * 9000);
      const pin = Math.floor(1000 + Math.random() * 9000).toString(); // 4-digit PIN
      const timestamp = new Date();
      
      // Exact alignment with resetDatabase schema:
      // ['Order ID', 'Customer Name', 'Phone', 'Area', 'Landmark', 'Item', 'Status', 'Rider ID', 'PIN', 'Timestamp']
      ordersSheet.appendRow([
        orderId, 
        data.customerName || '', 
        data.phone || '', 
        data.area || '', 
        data.landmark || '', 
        data.item || '', 
        'LOGGED', 
        'Unassigned', 
        pin, 
        timestamp
      ]);
      
      logEvent(orderId, 'Retailer', 'CREATE', '', 'LOGGED');
      
      // Return both orderId AND pin to the frontend
      return ContentService.createTextOutput(JSON.stringify({ success: true, orderId: orderId, pin: pin }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * DATABASE RESET & SEED SCRIPT
 */
function resetDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Reset Orders Tab
  let ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet('Orders');
  } else {
    ordersSheet.clear();
  }
  ordersSheet.appendRow([
    'Order ID', 'Customer Name', 'Phone', 'Area', 'Landmark', 'Item', 'Status', 'Rider ID', 'PIN', 'Timestamp'
  ]);
  ordersSheet.getRange("1:1").setFontWeight("bold").setBackground("#D9EAD3");

  // 2. Reset Events Tab
  let eventsSheet = ss.getSheetByName('Events');
  if (!eventsSheet) {
    eventsSheet = ss.insertSheet('Events');
  } else {
    eventsSheet.clear();
  }
  eventsSheet.appendRow([
    'Event ID', 'Order ID', 'Timestamp', 'Actor', 'Action', 'Old Status', 'New Status', 'Device', 'Sync Status'
  ]);
  eventsSheet.getRange("1:1").setFontWeight("bold").setBackground("#FFF2CC");

  // 3. Reset Riders Tab (Seed Data)
  let ridersSheet = ss.getSheetByName('Riders');
  if (!ridersSheet) {
    ridersSheet = ss.insertSheet('Riders');
  } else {
    ridersSheet.clear();
  }
  ridersSheet.appendRow(['Rider ID', 'Name', 'Phone', 'Vehicle', 'Status']);
  ridersSheet.appendRow(['Rider-1', 'John Doe', '0700000001', 'Motorbike', 'ACTIVE']);
  ridersSheet.appendRow(['Rider-2', 'Jane Smith', '0700000002', 'Bicycle', 'ACTIVE']);
  ridersSheet.getRange("1:1").setFontWeight("bold").setBackground("#C9DAF8");

  Logger.log("Database reset complete. Fresh schema headers and seed data applied!");
}