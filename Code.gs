/**
 * Routing logic: Serves Admin page, Doctor page, or the Patient Index page
 * based on the URL parameter (e.g. your-web-app-url?page=admin)
 */
function doGet(e) {
  const userAgent = e.userInterface ? "" : "mobile-detect"

  if (e.parameter.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
        .setTitle('PolyClinic Admin Portal')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
 
  if (e.parameter.page === 'doctor') {
    return HtmlService.createHtmlOutputFromFile('Doctor')
        .setTitle('PolyClinic Practitioner Portal')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
 
   if (e.parameter.page === 'patient') {
    return HtmlService.createHtmlOutputFromFile('Patient')
        .setTitle('PolyClinic Patient Portal')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('PolyClinic Appointment Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Run this function once to reset/build the exact columns.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Patient Table
  let patientSheet = ss.getSheetByName('patient');
  if (!patientSheet) {
    patientSheet = ss.insertSheet('patient');
    patientSheet.appendRow(['patient_id', 'first_name', 'last_name', 'mobile', 'email']);
    patientSheet.getRange("A1:E1").setFontWeight("bold");
    patientSheet.autoResizeColumns(1, 5);
  }

  // 2. Setup Practitioner Table
  let pracSheet = ss.getSheetByName('practitioner');
  if (!pracSheet) {
    pracSheet = ss.insertSheet('practitioner');
    pracSheet.appendRow(['practitioner_id', 'first_name', 'last_name', 'credential', 'services']);
    pracSheet.getRange("A1:E1").setFontWeight("bold");
    pracSheet.autoResizeColumns(1, 5);
  }

  // 3. Setup Schedule Table 
  let schedSheet = ss.getSheetByName('schedule');
  if (!schedSheet) {
    schedSheet = ss.insertSheet('schedule');
  }
  // Added 'loc_id' to align schedule slots to specific locations
  schedSheet.getRange("A1:F1").setValues([['schedule_id', 'practitioner_id', 'day_of_week', 'start_time', 'end_time', 'loc_id']]);
  schedSheet.getRange("A1:F1").setFontWeight("bold");
  
  // 4. Setup Appointment Table
  let apptSheet = ss.getSheetByName('appointment');
  if (!apptSheet) {
    apptSheet = ss.insertSheet('appointment');
  }
  // Updated Schema to include purpose (N), loc_id (O), and location (P)
  apptSheet.getRange("A1:P1").setValues([[
    'reference_id', 'patient_id', 'patient_name', 'date_submitted', 'apt_date', 'apt_start_time', 'apt_end_time',
    'service_category', 'practitioner', 'status', 'email_notification',
    'modification_timestamp', 'accept_decline_reason', 'purpose', 'loc_id', 'location'
  ]]);
  apptSheet.getRange("A1:P1").setFontWeight("bold");

  // 5. Setup Time Off Table
  let timeOffSheet = ss.getSheetByName('time_off');
  if (!timeOffSheet) {
    timeOffSheet = ss.insertSheet('time_off');
    timeOffSheet.appendRow(['time_off_id', 'practitioner_id', 'date', 'start_time', 'end_time', 'reason']);
    timeOffSheet.getRange("A1:F1").setFontWeight("bold");
    timeOffSheet.autoResizeColumns(1, 6);
  }

  // 6. Setup Location Table (NEW)
  let locSheet = ss.getSheetByName('location');
  if (!locSheet) {
    locSheet = ss.insertSheet('location');
    locSheet.appendRow(['loc_id', 'location_name', 'loc_address']);
    locSheet.getRange("A1:C1").setFontWeight("bold");
    locSheet.autoResizeColumns(1, 3);
  }
}

/**
 * Patient Portal: Fetch available slots
 */
function getSlotsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- NEW: Fetch Location Data ---
    const locSheet = ss.getSheetByName('location');
    const locations = {};
    const locList = [];
    if (locSheet) {
      const locData = locSheet.getDataRange().getValues();
      locData.shift(); 
      locData.forEach(row => {
        if (row[0]) {
          locations[String(row[0]).trim()] = {
            name: row[1],
            address: row[2]
          };
          locList.push({ id: String(row[0]).trim(), name: row[1], address: row[2] });
        }
      });
    }
    // --------------------------------

    const pracSheet = ss.getSheetByName('practitioner');
    const pracData = pracSheet.getDataRange().getValues();
    pracData.shift();

    const practitioners = {};
    const lobs = new Set();
    const pracNames = new Set();

    pracData.forEach(row => {
      if (!row[0]) return;
      const fullName = `${row[1]} ${row[2]}, ${row[3]}`;
      practitioners[row[0]] = { name: fullName, lob: row[4] };
      lobs.add(row[4]);
      pracNames.add(fullName);
    });

    const schedSheet = ss.getSheetByName('schedule');
    const schedData = schedSheet.getDataRange().getValues();
    schedData.shift();

    const rulesByDay = { 'Sunday': [], 'Monday': [], 'Tuesday': [], 'Wednesday': [], 'Thursday': [], 'Friday': [], 'Saturday': [] };

    schedData.forEach(row => {
      if (!row[2]) return;
      let startHour = row[3] instanceof Date ? row[3].getHours() : parseInt(String(row[3]).split(':')[0], 10) || 8;
      let endHour = row[4] instanceof Date ? row[4].getHours() : parseInt(String(row[4]).split(':')[0], 10) || 17;
      
      let scheduleLocId = row[5] ? String(row[5]).trim() : ''; 
      rulesByDay[row[2]].push({ practitioner_id: row[1], start: startHour, end: endHour, loc_id: scheduleLocId });
    });

    const timeOffSheet = ss.getSheetByName('time_off');
    const timeOffRules = [];
    if (timeOffSheet) {
      const timeOffData = timeOffSheet.getDataRange().getValues();
      timeOffData.shift(); 
      timeOffData.forEach(row => {
        if (!row[1]) return; 
        let offDateStr = '';
        if (row[2]) {
          offDateStr = row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(row[2]).trim();
        }
        function parseHour(val) {
          if (val === null || val === undefined || val === '') return null;
          if (val instanceof Date) return val.getHours();
          const str = String(val).trim().toUpperCase();
          let match = str.match(/(\d+)/);
          if (!match) return null;
          let h = parseInt(match[1], 10);
          if (str.includes('PM') && h < 12) h += 12;
          if (str.includes('AM') && h === 12) h = 0;
          return h;
        }
        let sHour = parseHour(row[3]);
        let eHour = parseHour(row[4]);
        if (sHour !== null && eHour !== null) {
          timeOffRules.push({ practitioner_id: String(row[1]).trim(), dateStr: offDateStr, startHour: sHour, endHour: eHour });
        }
      });
    }

    const apptSheet = ss.getSheetByName('appointment');
    const apptData = apptSheet.getDataRange().getValues();
    apptData.shift();

    const bookedSet = new Set();
    apptData.forEach(row => {
      if (!row[4]) return; 
      let dateStr = row[4] instanceof Date ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : row[4];
      const startTime = row[5] instanceof Date ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'hh:mm a') : row[5];
      const pracName = row[8]; 
      const status = row[9]; 

      if (status !== 'Declined' && status !== 'Cancelled' && status !== 'Rejected') {
        bookedSet.add(`${dateStr}|${startTime}|${pracName}`);
      }
    });

    const slots = [];
    const today = new Date();

    for (let i = 0; i < 60; i++) {
      const d = new Date(today);
      d.setDate(today.getDate() + i);
      const dayName = Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEEE');
      const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      const dayRules = rulesByDay[dayName] || [];
      dayRules.forEach(rule => {
        const prac = practitioners[rule.practitioner_id];
        if (!prac) return;
        const locName = locations[rule.loc_id] ? locations[rule.loc_id].name : 'Default Location';
        const locAddress = locations[rule.loc_id] ? locations[rule.loc_id].address : '';

        for (let hour = rule.start; hour < rule.end; hour++) {
          let isOff = false;
          for (let t = 0; t < timeOffRules.length; t++) {
            let to = timeOffRules[t];
            if (to.practitioner_id === rule.practitioner_id && (!to.dateStr || to.dateStr === dateStr)) {
              if (hour >= to.startHour && hour < to.endHour) {
                isOff = true;
                break;
              }
            }
          }
          if (isOff) continue;

          const formattedStart = (hour > 12 ? String(hour - 12).padStart(2, '0') : String(hour).padStart(2, '0')) + ':00 ' + (hour >= 12 ? 'PM' : 'AM');
          const formattedEnd = ((hour + 1) > 12 ? String(hour + 1 - 12).padStart(2, '0') : String(hour + 1).padStart(2, '0')) + ':00 ' + ((hour + 1) >= 12 ? 'PM' : 'AM');
          const timeSlot = `${formattedStart} - ${formattedEnd}`;
          const isBooked = bookedSet.has(`${dateStr}|${formattedStart}|${prac.name}`);

          slots.push({
            practitionerId: rule.practitioner_id,
            practitionerName: prac.name,
            lob: prac.lob,
            loc_id: rule.loc_id,
            location: locName,
            locAddress: locAddress,
            date: dateStr,
            timeSlot: timeSlot,
            start_time: formattedStart,
            end_time: formattedEnd,
            status: isBooked ? 'Booked' : 'Available'
          });
        }
      });
    }

    return { 
      success: true, 
      data: slots, 
      specialties: Array.from(lobs).sort(), 
      doctors: Array.from(pracNames).sort(),
      locations: locList 
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Patient Portal: Submit booking request
 */
function bookSlot(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const apptSheet = ss.getSheetByName('appointment');
    
    const aData = apptSheet.getDataRange().getValues();
    const headers = aData[0];

    // Concurrency Check
    let isOverlapping = false;
    for (let i = 1; i < aData.length; i++) {
      let r = aData[i];
      let rDate = r[4] instanceof Date ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : r[4];
      let rStart = r[5] instanceof Date ? Utilities.formatDate(r[5], Session.getScriptTimeZone(), 'hh:mm a') : r[5];
      let rPrac = r[8];
      let rStatus = r[9];

      if (rDate === payload.date && rStart === payload.start_time && rPrac === payload.practitionerName) {
        if (rStatus !== 'Declined' && rStatus !== 'Cancelled' && rStatus !== 'Rejected') {
           isOverlapping = true;
           break;
        }
      }
    }

    // Auto-create or fetch patient ID
    const pSheet = ss.getSheetByName('patient');
    const pData = pSheet.getDataRange().getValues();
    let patientId = '';
    
    for (let i = 1; i < pData.length; i++) {
       if (String(pData[i][3]).trim() === payload.contactNumber || String(pData[i][4]).trim().toLowerCase() === payload.emailAddress.toLowerCase()) {
         patientId = pData[i][0];
         break;
       }
    }

    if (!patientId) {
      patientId = 'PT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM') + '-' + String(pData.length).padStart(4, '0');
      pSheet.appendRow([patientId, payload.firstName, payload.lastName, payload.contactNumber, payload.emailAddress]);
    }

    const referenceId = 'REQ-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + String(aData.length).padStart(4, '0');

    // Appending using the proper 16-column structure
    apptSheet.appendRow([
      referenceId,
      patientId,
      `${payload.firstName} ${payload.lastName}`,
      new Date(),
      payload.date,
      payload.start_time,
      payload.end_time,
      payload.lob,
      payload.practitionerName,
      'Pending',
      'Pending',
      '',
      '',
      payload.purpose,        // Column N
      payload.loc_id || '',   // Column O
      payload.location || ''  // Column P
    ]);

    return { success: true, referenceId: referenceId, isOverlapping: isOverlapping };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * ADMIN PORTAL: Process an Accept or Decline action
 */
function processAdminAction(payload) {
  try {
    const { referenceId, action, reason } = payload;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const apptSheet = ss.getSheetByName('appointment');
    const patientSheet = ss.getSheetByName('patient');
   
    const apptData = apptSheet.getDataRange().getValues();
    let rowIndex = -1;
    let apptRow = null;
   
    for (let i = 1; i < apptData.length; i++) {
      if (apptData[i][0] === referenceId) {
        rowIndex = i + 1;
        apptRow = apptData[i];
        break;
      }
    }
   
    if (rowIndex === -1) throw new Error("Appointment not found.");
   
    const newStatus = action === 'accept' ? 'Booked' : 'Declined';
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm:ss');
   
    apptSheet.getRange(rowIndex, 10).setValue(newStatus); // Status is column J
    apptSheet.getRange(rowIndex, 12).setValue(timestamp); // Modification Timestamp is column L
    if (reason) apptSheet.getRange(rowIndex, 13).setValue(reason); // Reason is column M
   
    // Fetch Patient Email
    const patientId = apptRow[1];
    const pData = patientSheet.getDataRange().getValues();
    let pEmail = '';
    let pName = apptRow[2];
   
    for(let i = 1; i < pData.length; i++) {
      if(pData[i][0] === patientId) {
        pEmail = pData[i][4];
        break;
      }
    }
   
    // Send Email Notification
    if (pEmail) {
      let subject, body;
      let formattedDate = apptRow[4] instanceof Date ? Utilities.formatDate(apptRow[4], Session.getScriptTimeZone(), 'MMMM dd, yyyy') : apptRow[4];
      let formattedStartTime = apptRow[5] instanceof Date ? Utilities.formatDate(apptRow[5], Session.getScriptTimeZone(), 'hh:mm a') : apptRow[5];
      let formattedEndTime = apptRow[6] instanceof Date ? Utilities.formatDate(apptRow[6], Session.getScriptTimeZone(), 'hh:mm a') : apptRow[6];

      if (action === 'accept') {
        subject = `Appointment Confirmed - PolyClinic (${referenceId})`;
        body = `Dear ${pName},<br><br>Good news! Your appointment request has been <b>Confirmed</b>.<br><br><b>Details:</b><br>Date: ${formattedDate}<br>Time: ${formattedStartTime} - ${formattedEndTime}<br>Practitioner: ${apptRow[8]}<br><br>Please arrive 15 minutes early.<br><br>Thank you,<br>PolyClinic Administration`;
      } else {
        subject = `Appointment Request Declined - PolyClinic (${referenceId})`;
        body = `Dear ${pName},<br><br>We regret to inform you that your appointment request for ${formattedDate} at ${formattedStartTime} - ${formattedEndTime} with ${apptRow[8]} has been <b>Declined</b>.<br><br><b>Reason from Admin:</b><br>${reason}<br><br>Please feel free to submit a new request for a different date or time.<br><br>Thank you,<br>PolyClinic Administration`;
      }
     
      GmailApp.sendEmail(
        pEmail,
        subject,
        '',
        {
          name: "Poly Clinic Scheduler",
          htmlBody: body }
          );

      apptSheet.getRange(rowIndex, 11).setValue('Yes');
    }
   
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * DOCTOR PORTAL: Alias for processAdminAction to securely handle Accept/Decline
 * requests submitted directly from the Practitioner portal.
 */
function processDoctorAction(payload) {
  return processAdminAction(payload);
}

/**
 * ADMIN PORTAL: Fetch all appointments
 */
function getAdminData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
   
    const patientSheet = ss.getSheetByName('patient');
    const pData = patientSheet.getDataRange().getValues();
    const patientMap = {};
    for (let i = 1; i < pData.length; i++) {
      patientMap[pData[i][0]] = { email: pData[i][4], phone: pData[i][3] };
    }


    const apptSheet = ss.getSheetByName('appointment');
    const aData = apptSheet.getDataRange().getValues();
    aData.shift();
   
    const appointments = aData.map(row => {
      let submitDate = row[3] instanceof Date ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm') : row[3];
      let aptDate = row[4] instanceof Date ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'MMM dd, yyyy') : row[4];
      let aptStartTime = row[5] instanceof Date ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'hh:mm a') : row[5];
      let aptEndTime = row[6] instanceof Date ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'hh:mm a') : row[6];


      const patInfo = patientMap[row[1]] || { email: 'N/A', phone: 'N/A' };
     
      return {
        refId: row[0],
        patientName: row[2],
        patientEmail: patInfo.email,
        dateSubmitted: submitDate,
        aptDate: aptDate,
        time: `${aptStartTime} - ${aptEndTime}`,
        service: row[7],
        practitioner: row[8],
        status: row[9],
        emailNotif: row[10],
        patientPurpose: row[13]
      };
    });


    appointments.reverse();


    return { success: true, appointments: appointments };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}/**
 * NEW: Practitioner Portal - Fetch list of practitioners for the dropdown
 */
function getPractitionersList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pracSheet = ss.getSheetByName('practitioner');
    const pracData = pracSheet.getDataRange().getValues();
    pracData.shift(); // Remove headers
   
    const names = new Set();
    pracData.forEach(row => {
      if (!row[0]) return;
      names.add(`${row[1]} ${row[2]}, ${row[3]}`);
    });
   
    return { success: true, data: Array.from(names).sort() };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}


/**
 * NEW: Practitioner Portal - Fetch appointments specifically for one doctor
 */
function getDoctorAppointments(doctorName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
   
    const patientSheet = ss.getSheetByName('patient');
    const pData = patientSheet.getDataRange().getValues();
    const patientMap = {};
    for (let i = 1; i < pData.length; i++) {
      patientMap[pData[i][0]] = { email: pData[i][4], phone: pData[i][3] };
    }


    const apptSheet = ss.getSheetByName('appointment');
    const aData = apptSheet.getDataRange().getValues();
    aData.shift();
   
    const appointments = [];
   
    aData.forEach(row => {
      // Filter by the requested doctor
      if (row[8] !== doctorName) return;
     
      let aptDate = row[4] instanceof Date ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'MMM dd, yyyy') : row[4];
      let aptStartTime = row[5] instanceof Date ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'hh:mm a') : row[5];
      let aptEndTime = row[6] instanceof Date ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'hh:mm a') : row[6];


      const patInfo = patientMap[row[1]] || { email: 'N/A', phone: 'N/A' };
     
      appointments.push({
        refId: row[0],
        patientName: row[2],
        patientEmail: patInfo.email,
        patientPhone: patInfo.phone,
        aptDate: aptDate,
        time: `${aptStartTime} - ${aptEndTime}`,
        service: row[7],
        status: row[9],
        patientPurpose: row[13]
      });
    });


    // Sort so earliest dates are first
    appointments.sort((a, b) => new Date(a.aptDate) - new Date(b.aptDate));


    return { success: true, appointments: appointments };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}


