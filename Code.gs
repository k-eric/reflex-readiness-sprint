/**
 * REFLEX BACKEND LOGIC - FULLY INTEGRATED & ALIGNED
 */

// Helper to append a new event to the audit trail
function logEvent(orderId, actor, action, oldStatus, newStatus, device = 'Web', syncStatus = 'Synced') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = ss.getSheetByName('Events');
  const eventId = 'EV-' + new Date().getTime();
  const timestamp = new Date();
  
  eventsSheet.appendRow([eventId, orderId, timestamp, actor, action, oldStatus, newStatus, device, syncStatus]);
}

// State transition validator, status updater, & rider re-assignment handler
function updateOrderStatus(orderId, newStatus, actorRole, riderId = null, pin = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  const data = ordersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      const currentStatus = data[i][6]; // Column 7: Status (Index 6)
      const storedPin = String(data[i][8]); // Column 9: PIN (Index 8)
      
      // 1. RE-ASSIGNMENT LOGIC (ASSIGNED -> ASSIGNED)
      if (currentStatus === 'ASSIGNED' && newStatus === 'ASSIGNED') {
        if (riderId) {
          ordersSheet.getRange(i + 1, 8).setValue(riderId); // Update Rider ID
        }
        logEvent(orderId, actorRole, 'REASSIGN_RIDER', currentStatus, newStatus);
        return { success: true, message: `Reassigned order ${orderId} to ${riderId}` };
      }

      // 2. PIN VERIFICATION FOR DELIVERED STATUS
      if (newStatus === 'DELIVERED') {
        if (String(pin).trim() !== storedPin.trim()) {
          return { success: false, message: 'Invalid PIN! Handoff verification failed.' };
        }
      }

      // 3. STANDARD TRANSITIONS (LOGGED -> ASSIGNED, ASSIGNED -> PICKED UP, PICKED UP -> DELIVERED)
      const validTransitions = {
        'LOGGED': 'ASSIGNED',
        'ASSIGNED': 'PICKED UP',
        'PICKED UP': 'DELIVERED'
      };

      if (validTransitions[currentStatus] === newStatus) {
        ordersSheet.getRange(i + 1, 7).setValue(newStatus); // Update Status
        
        // Update Rider ID if assigning initial rider
        if (riderId) {
          ordersSheet.getRange(i + 1, 8).setValue(riderId);
        }

        logEvent(orderId, actorRole, 'STATUS_UPDATE', currentStatus, newStatus);
        return { success: true, message: `Status updated from ${currentStatus} to ${newStatus}` };
      } else {
        return { success: false, message: `Invalid transition from ${currentStatus} to ${newStatus}` };
      }
    }
  }
  return { success: false, message: 'Order ID not found' };
}

// Expose GET endpoint for fetching orders safely
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Orders');
    
    if (!ordersSheet) {
      return ContentService.createTextOutput(JSON.stringify({ success: true, orders: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = ordersSheet.getDataRange().getValues();
    
    // If only headers exist or sheet is empty
    if (!data || data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ success: true, orders: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const orders = [];
    for (let i = 1; i < data.length; i++) {
      // Only parse non-empty rows
      if (data[i][0]) {
        orders.push({
          orderId: String(data[i][0] || ''),
          customerName: String(data[i][1] || ''),
          phone: String(data[i][2] || ''),
          area: String(data[i][3] || ''),
          landmark: String(data[i][4] || ''),
          item: String(data[i][5] || ''),
          status: String(data[i][6] || 'LOGGED'),
          riderId: String(data[i][7] || 'Unassigned'),
          pin: String(data[i][8] || ''),
          timestamp: String(data[i][9] || '')
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, orders: orders }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // Return error as JSON so CORS doesn't trigger
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Expose POST endpoint for creating AND updating orders
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. CREATE ORDER (LOGGED)
    if (data.action === 'CREATE_ORDER') {
      const ordersSheet = ss.getSheetByName('Orders');
      const orderId = 'RX-' + Math.floor(1000 + Math.random() * 9000);
      const pin = Math.floor(1000 + Math.random() * 9000).toString(); // 4-digit PIN
      const timestamp = new Date();
      
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
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, orderId: orderId, pin: pin }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 2. UPDATE ORDER STATUS (ASSIGN, REASSIGN, PICKUP, DELIVER)
    if (data.action === 'UPDATE_STATUS') {
      const result = updateOrderStatus(
        data.orderId, 
        data.newStatus, 
        data.actorRole || 'System', 
        data.riderId, 
        data.pin
      );
      return ContentService.createTextOutput(JSON.stringify(result))
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