/**
 * REFLEX BACKEND LOGIC - ALL PHASES INTEGRATED (1, 2, 3 & 4)
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

// Helper to fetch full list of registered riders with vehicle & range info
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
        status: String(data[i][5] || 'ACTIVE')
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
      const currentStatus = data[i][8];
      const orderDistance = Number(data[i][17]) || 0;

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

      if ((currentStatus === 'ASSIGNED' || currentStatus === 'FAILED') && newStatus === 'ASSIGNED') {
        if (riderId) ordersSheet.getRange(i + 1, 10).setValue(riderId);
        ordersSheet.getRange(i + 1, 9).setValue(newStatus);
        logEvent(orderId, actorRole, 'REASSIGN_RIDER', currentStatus, newStatus);
        return { success: true, message: `Reassigned order ${orderId} to ${riderId}` };
      }

      if (newStatus === 'FAILED') {
        ordersSheet.getRange(i + 1, 9).setValue('FAILED');
        if (failureReason) ordersSheet.getRange(i + 1, 13).setValue(failureReason);
        logEvent(orderId, actorRole, 'DELIVERY_FAILED', currentStatus, 'FAILED');
        return { success: true, message: `Order ${orderId} marked as FAILED (${failureReason})` };
      }

      if (currentStatus === 'FAILED' && newStatus === 'LOGGED') {
        ordersSheet.getRange(i + 1, 9).setValue('LOGGED');
        ordersSheet.getRange(i + 1, 10).setValue('Unassigned');
        ordersSheet.getRange(i + 1, 13).setValue('');
        logEvent(orderId, actorRole, 'RESET_ORDER', currentStatus, 'LOGGED');
        return { success: true, message: `Order ${orderId} reset back to LOGGED` };
      }

      if (newStatus === 'DELIVERED') {
        const storedPin = String(data[i][10]);
        if (String(pin).trim() !== storedPin.trim()) {
          return { success: false, message: 'Invalid PIN! Handoff verification failed.' };
        }
      }

      const validTransitions = {
        'LOGGED': 'ASSIGNED',
        'ASSIGNED': 'PICKED UP',
        'PICKED UP': 'DELIVERED'
      };

      if (validTransitions[currentStatus] === newStatus) {
        ordersSheet.getRange(i + 1, 9).setValue(newStatus);
        if (riderId) ordersSheet.getRange(i + 1, 10).setValue(riderId);
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
          customerName: String(data[i][2] || ''),
          phone: String(data[i][3] || ''),
          area: String(data[i][4] || ''),
          landmark: String(data[i][5] || ''),
          item: String(data[i][6] || ''),
          itemModel: String(data[i][7] || 'N/A'),
          status: String(data[i][8] || 'LOGGED'),
          riderId: String(data[i][9] || 'Unassigned'),
          pin: String(data[i][10] || ''),
          timestamp: String(data[i][11] || ''),
          failureReason: String(data[i][12] || ''),
          deliveryType: String(data[i][13] || 'Immediate'),
          isUrgent: Boolean(data[i][14]),
          deliveryDate: String(data[i][15] || ''),
          timeSlot: String(data[i][16] || ''),
          estimatedDistance: Number(data[i][17]) || 0,
          isOutOfZone: Boolean(data[i][18])
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
        data.customerName || '', 
        data.phone || '', 
        data.area || '', 
        data.landmark || '', 
        data.item || '', 
        data.itemModel || 'N/A',
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
    'Order ID', 'Shop Name', 'Customer Name', 'Phone', 'Area', 'Landmark', 'Item', 'Item Model', 
    'Status', 'Rider ID', 'PIN', 'Timestamp', 'Failure Reason', 
    'Delivery Type', 'Is Urgent', 'Delivery Date', 'Time Slot', 'Est Distance (km)', 'Out of Zone'
  ]);
  ordersSheet.getRange("1:1").setFontWeight("bold").setBackground("#D9EAD3");

  let eventsSheet = ss.getSheetByName('Events');
  if (!eventsSheet) eventsSheet = ss.insertSheet('Events');
  else eventsSheet.clear();

  eventsSheet.appendRow(['Event ID', 'Order ID', 'Timestamp', 'Actor', 'Action', 'Old Status', 'New Status', 'Device', 'Sync Status']);
  eventsSheet.getRange("1:1").setFontWeight("bold").setBackground("#FFF2CC");

  // Reset Riders Tab with Fleet Vehicle & Max Range Security Capacity
  let ridersSheet = ss.getSheetByName('Riders');
  if (!ridersSheet) ridersSheet = ss.insertSheet('Riders');
  else ridersSheet.clear();

  ridersSheet.appendRow(['Rider ID', 'Name', 'Phone', 'Vehicle Type', 'Max Range (km)', 'Status']);
  ridersSheet.appendRow(['Rider-1', 'John Doe', '0700000001', 'Bicycle', 10, 'ACTIVE']);
  ridersSheet.appendRow(['Rider-2', 'Jane Smith', '0700000002', 'Boda Boda', 25, 'ACTIVE']);
  ridersSheet.appendRow(['Rider-3', 'Mark Kamau', '0700000003', 'Pickup Van', 150, 'ACTIVE']);
  ridersSheet.getRange("1:1").setFontWeight("bold").setBackground("#C9DAF8");

  Logger.log("Database reset complete. All 4 Phases now active with Rider Fleet Capacity Controls!");
}