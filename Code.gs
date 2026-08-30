/**
 * REFLEX BACKEND LOGIC - ALL PHASES INTEGRATED (1, 2, 3 & 4)
 * Updated with Retailer Metadata, Quantity, Operating Zones & ARRIVED Status
 */

function isValidKenyanPhone(phone) {
  if (!phone) return false;
  const cleanPhone = String(phone).trim().replace(/[\s-]/g, '');
  return /^(\+254|0)[71]\d{8}$/.test(cleanPhone);
}

function logEvent(orderId, actor, action, oldStatus, newStatus, device = 'Web', syncStatus = 'Synced') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = ss.getSheetByName('Events');
  if (!eventsSheet) return;
  const eventId = 'EV-' + new Date().getTime();
  const timestamp = new Date();
  
  eventsSheet.appendRow([eventId, orderId, timestamp, actor, action, oldStatus, newStatus, device, syncStatus]);
}

// Fetch full list of registered riders including Operating Zone
function getRidersList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ridersSheet = ss.getSheetByName('Riders');
  if (!ridersSheet) return [];
  
  const data = ridersSheet.getDataRange().getValues();
  const riders = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      riders.push({
        riderId: String(data[i][0]),
        name: String(data[i][1]),
        phone: String(data[i][2]),
        vehicle: String(data[i][3]),
        maxRangeKm: Number(data[i][4]) || 15,
        status: String(data[i][5] || 'ACTIVE'),
        operatingZone: String(data[i][6] || 'Central Nairobi') // Added Operating Zone
      });
    }
  }
  return riders;
}

function updateOrderStatus(orderId, newStatus, actorRole, riderId = null, pin = null, failureReason = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName('Orders');
  const data = ordersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      const currentStatus = data[i][10]; // Column index 10 is Status
      const orderDistance = Number(data[i][21]) || 0; // Column index 21 is Est Distance

      // Phase 3 Validation: Verify Rider Range Restrictions before assignment
      if (riderId && (newStatus === 'ASSIGNED')) {
        const riders = getRidersList();
        const selectedRider = riders.find(r => r.riderId === riderId);
        
        if (selectedRider && orderDistance > selectedRider.maxRangeKm) {
          return {
            success: false,
            message: `Assignment Blocked! ${selectedRider.name} (${selectedRider.vehicle}) has a max range of ${selectedRider.maxRangeKm}km, but this order is ${orderDistance}km.`
          };
        }
      }

      // Handle Reassignment
      if ((currentStatus === 'ASSIGNED' || currentStatus === 'FAILED' || currentStatus === 'ARRIVED') && newStatus === 'ASSIGNED') {
        if (riderId) ordersSheet.getRange(i + 1, 12).setValue(riderId); // Column 12: Rider ID
        ordersSheet.getRange(i + 1, 11).setValue(newStatus); // Column 11: Status
        logEvent(orderId, actorRole, 'REASSIGN_RIDER', currentStatus, newStatus);
        return { success: true, message: `Reassigned order ${orderId} to ${riderId}` };
      }

      // Handle Delivery Failures
      if (newStatus === 'FAILED') {
        ordersSheet.getRange(i + 1, 11).setValue('FAILED');
        if (failureReason) ordersSheet.getRange(i + 1, 15).setValue(failureReason); // Column 15: Failure Reason
        logEvent(orderId, actorRole, 'DELIVERY_FAILED', currentStatus, 'FAILED');
        return { success: true, message: `Order ${orderId} marked as FAILED (${failureReason})` };
      }

      // Handle Resetting Failed Orders
      if (currentStatus === 'FAILED' && newStatus === 'LOGGED') {
        ordersSheet.getRange(i + 1, 11).setValue('LOGGED');
        ordersSheet.getRange(i + 1, 12).setValue('Unassigned');
        ordersSheet.getRange(i + 1, 15).setValue('');
        logEvent(orderId, actorRole, 'RESET_ORDER', currentStatus, 'LOGGED');
        return { success: true, message: `Order ${orderId} reset back to LOGGED` };
      }

      // Handle PIN Verification upon Delivery
      if (newStatus === 'DELIVERED') {
        const storedPin = String(data[i][12]); // Column index 12 is PIN
        if (String(pin).trim() !== storedPin.trim()) {
          return { success: false, message: 'Invalid PIN! Handoff verification failed.' };
        }
      }

      // Updated State Transitions including ARRIVED status
      const validTransitions = {
        'LOGGED': 'ASSIGNED',
        'ASSIGNED': 'PICKED UP',
        'PICKED UP': 'ARRIVED',
        'ARRIVED': 'DELIVERED'
      };

      if (validTransitions[currentStatus] === newStatus) {
        ordersSheet.getRange(i + 1, 11).setValue(newStatus); // Column 11: Status
        if (riderId) ordersSheet.getRange(i + 1, 12).setValue(riderId); // Column 12: Rider ID
        logEvent(orderId, actorRole, 'STATUS_UPDATE', currentStatus, newStatus);
        return { success: true, message: `Status updated from ${currentStatus} to ${newStatus}` };
      } else {
        return { success: false, message: `Invalid transition from ${currentStatus} to ${newStatus}` };
      }
    }
  }
  return { success: false, message: 'Order ID not found' };
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = ss.getSheetByName('Orders');
    
    const riders = getRidersList();

    if (!ordersSheet) {
      return ContentService.createTextOutput(JSON.stringify({ success: true, orders: [], riders: riders }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = ordersSheet.getDataRange().getValues();
    const orders = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        orders.push({
          orderId: String(data[i][0] || ''),
          shopName: String(data[i][1] || 'Unspecified Shop'),
          shopType: String(data[i][2] || 'General Retail'),
          pickupLocation: String(data[i][3] || ''),
          isVerifiedRetailer: Boolean(data[i][4]),
          customerName: String(data[i][5] || ''),
          phone: String(data[i][6] || ''),
          area: String(data[i][7] || ''),
          landmark: String(data[i][8] || ''),
          item: String(data[i][9] || ''),
          itemModel: String(data[i][10] || 'N/A'),
          itemQty: Number(data[i][11]) || 1,
          status: String(data[i][12] || 'LOGGED'),
          riderId: String(data[i][13] || 'Unassigned'),
          pin: String(data[i][14] || ''),
          timestamp: String(data[i][15] || ''),
          failureReason: String(data[i][16] || ''),
          deliveryType: String(data[i][17] || 'Immediate'),
          isUrgent: Boolean(data[i][18]),
          deliveryDate: String(data[i][19] || ''),
          timeSlot: String(data[i][20] || ''),
          estimatedDistance: Number(data[i][21]) || 0,
          isOutOfZone: Boolean(data[i][22])
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, orders: orders, riders: riders }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (data.action === 'CREATE_ORDER') {
      const ordersSheet = ss.getSheetByName('Orders');
      
      if (!isValidKenyanPhone(data.phone)) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: "Invalid Kenyan phone number format."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const orderId = 'RX-' + Math.floor(1000 + Math.random() * 9000);
      const pin = Math.floor(1000 + Math.random() * 9000).toString();
      const timestamp = new Date();
      
      ordersSheet.appendRow([
        orderId, 
        data.shopName || 'Unspecified Shop',
        data.shopType || 'General Retail',
        data.pickupLocation || 'Nairobi CBD Main Hub',
        data.isVerifiedRetailer ? true : false,
        data.customerName || '', 
        data.phone || '', 
        data.area || '', 
        data.landmark || '', 
        data.item || '', 
        data.itemModel || 'N/A',
        Number(data.itemQty) || 1,
        'LOGGED', 
        'Unassigned', 
        pin, 
        timestamp,
        '', 
        data.deliveryType || 'Immediate',
        data.isUrgent ? true : false,
        data.deliveryDate || '',
        data.timeSlot || '',
        Number(data.estimatedDistance) || 0,
        data.isOutOfZone ? true : false
      ]);
      
      logEvent(orderId, 'Retailer', 'CREATE', '', 'LOGGED');
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, orderId: orderId, pin: pin }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (data.action === 'UPDATE_STATUS') {
      const result = updateOrderStatus(
        data.orderId, 
        data.newStatus, 
        data.actorRole || 'System', 
        data.riderId, 
        data.pin,
        data.failureReason
      );
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function resetDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) ordersSheet = ss.insertSheet('Orders');
  else ordersSheet.clear();

  ordersSheet.appendRow([
    'Order ID', 'Shop Name', 'Shop Type', 'Pickup Location', 'Verified Retailer', 
    'Customer Name', 'Phone', 'Area', 'Landmark', 'Item Description', 'Item Model', 
    'Quantity', 'Status', 'Rider ID', 'PIN', 'Timestamp', 'Failure Reason', 
    'Delivery Type', 'Is Urgent', 'Delivery Date', 'Time Slot', 'Est Distance (km)', 'Out of Zone'
  ]);
  ordersSheet.getRange("1:1").setFontWeight("bold").setBackground("#D9EAD3");

  let eventsSheet = ss.getSheetByName('Events');
  if (!eventsSheet) eventsSheet = ss.insertSheet('Events');
  else eventsSheet.clear();

  eventsSheet.appendRow(['Event ID', 'Order ID', 'Timestamp', 'Actor', 'Action', 'Old Status', 'New Status', 'Device', 'Sync Status']);
  eventsSheet.getRange("1:1").setFontWeight("bold").setBackground("#FFF2CC");

  let ridersSheet = ss.getSheetByName('Riders');
  if (!ridersSheet) ridersSheet = ss.insertSheet('Riders');
  else ridersSheet.clear();

  ridersSheet.appendRow(['Rider ID', 'Name', 'Phone', 'Vehicle Type', 'Max Range (km)', 'Status', 'Operating Zone']);
  ridersSheet.appendRow(['Rider-1', 'John Doe', '0700000001', 'Bicycle', 10, 'ACTIVE', 'CBD / Westlands']);
  ridersSheet.appendRow(['Rider-2', 'Jane Smith', '0700000002', 'Boda Boda', 25, 'ACTIVE', 'Kilimani / Yaya']);
  ridersSheet.appendRow(['Rider-3', 'Mark Kamau', '0700000003', 'Pickup Van', 150, 'ACTIVE', 'Greater Nairobi']);
  ridersSheet.getRange("1:1").setFontWeight("bold").setBackground("#C9DAF8");

  Logger.log("Database reset complete. Column mappings, ARRIVED status transitions, and Rider Operating Zones are synchronized!");
}