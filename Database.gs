// --- IDENTIY & DIRECT LINE ACCESS CONTROL ---
// --- IDENTITY & DIRECT LINE ACCESS CONTROL ---
function getUserProfile() {
  const email = Session.getActiveUser().getEmail() || "Guest/Test User"; 
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User Roles');
    if (!sheet) return { email: email, role: 'Admin', scope: { division: [], group: [], department: [], section: [] }, l1Approver: "", l2Approver: "" }; 

    const data = sheet.getDataRange().getValues();
    data.shift(); 

    // Helper to split comma-separated items into clean array list
    const parseList = (str) => {
      if (!str) return [];
      return String(str)
        .split(',')
        .map(item => item.trim())
        .filter(item => item.length > 0);
    };

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) {
        const rawRole = String(data[i][1]).trim().toUpperCase();
        let normalizedRole = 'Standard';
        if (rawRole === 'ADMIN') normalizedRole = 'Admin';
        else if (rawRole === 'LEVEL 1 APPROVER') normalizedRole = 'Level 1 Approver';
        else if (rawRole === 'LEVEL 2 APPROVER') normalizedRole = 'Level 2 Approver';

        return {
          email: email,
          role: normalizedRole,
          scope: { 
            division: parseList(data[i][2]), 
            group: parseList(data[i][3]), 
            department: parseList(data[i][4]), 
            section: parseList(data[i][5]) 
          },
          l1Approver: data[i][6] ? String(data[i][6]).trim().toLowerCase() : "",
          l2Approver: data[i][7] ? String(data[i][7]).trim().toLowerCase() : ""
        };
      }
    }
    return { email: email, role: 'Standard', scope: { division: ['UNAUTHORIZED'], group: [], department: [], section: [] }, l1Approver: "", l2Approver: "" }; 
  } catch(e) {
    return { email: email, role: 'Standard', scope: { division: [], group: [], department: [], section: [] }, l1Approver: "", l2Approver: "" };
  }
}

// Helper to check multi-scope permissions (case-insensitive)
function isAllowedByScope(userScopeList, rowValue) {
  if (!userScopeList || userScopeList.length === 0) return true; // Empty scope = unrestricted at this level
  const target = String(rowValue || "").trim().toLowerCase();
  return userScopeList.some(scopeItem => scopeItem.toLowerCase() === target);
}

function getOrgChartData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    const orgSheet = ss.getSheetByName('OrgChart');
    if (!orgSheet) throw new Error("OrgChart sheet missing.");
    const orgDataRaw = orgSheet.getDataRange().getValues();
    orgDataRaw.shift(); 
    
    const orgChart = {};
    orgDataRaw.forEach(row => {
      const [div, grp, dept, sec] = row;
      if (!div) return;
      if (!orgChart[div]) orgChart[div] = {};
      if (!orgChart[div][grp]) orgChart[div][grp] = {};
      if (!orgChart[div][grp][dept]) orgChart[div][grp][dept] = [];
      if (sec && !orgChart[div][grp][dept].includes(sec)) orgChart[div][grp][dept].push(sec);
    });

    const refSheet = ss.getSheetByName('Reference Data');
    let locations = [];
    if (refSheet) {
      const refData = refSheet.getRange("A2:A").getValues();
      refData.forEach(row => { if (row[0]) locations.push(String(row[0]).trim()); });
    }

    const userProfile = getUserProfile();
    return { success: true, data: { orgChart, locations, userProfile } };
  } catch (e) { 
    return { success: false, message: "Init failed: " + e.message }; 
  }
}

function getNormalizedData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workload');
  const data = sheet.getDataRange().getValues();
  data.shift(); 

  const normalized = [];
  let currentBase = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowPos = String(row[0]).trim();
    const rowFte = String(row[16]).trim();
    
    let isFirstRow = false;
    if (i === 0) isFirstRow = true; 
    else if (rowFte !== "") isFirstRow = true; 
    else if (rowPos !== "" && rowPos !== currentBase.position) isFirstRow = true; 
    else if (rowPos !== "" && String(data[i-1][0]).trim() === "") isFirstRow = true; 

    if (isFirstRow) {
      currentBase = {
        position: rowPos || currentBase.position,
        fwa: String(row[1]).trim() || currentBase.fwa,
        netReportingHours: String(row[2]).trim() || currentBase.netReportingHours,
        division: String(row[3]).trim() || currentBase.division,
        group: String(row[4]).trim() || currentBase.group,
        department: String(row[5]).trim() || currentBase.department,
        section: String(row[6]).trim() || currentBase.section,
        employeeName: row[17] ? String(row[17]).trim() : "",
        workLocation: row[18] ? String(row[18]).trim() : "", 
        empType: row[19] ? String(row[19]).trim() : "",
        agencyName: row[20] ? String(row[20]).trim() : "", 
        empStatus: row[21] ? String(row[21]).trim() : "",
        startDate: row[22] ? String(row[22]).trim() : "", 
        endDate: row[23] ? String(row[23]).trim() : "",
        kpi1: row[24] ? String(row[24]).trim() : "", 
        kpi2: row[25] ? String(row[25]).trim() : "",
        kpi3: row[26] ? String(row[26]).trim() : "", 
        kpi4: row[27] ? String(row[27]).trim() : "",
        kpi5: row[28] ? String(row[28]).trim() : "",
        approvalStatus: row[29] ? String(row[29]).trim() : "Approved",
        remarks: row[30] ? String(row[30]).trim() : "",
        encoderEmail: row[31] ? String(row[31]).trim() : "",
        l1Target: row[32] ? String(row[32]).trim().toLowerCase() : "",
        l2Target: row[33] ? String(row[33]).trim().toLowerCase() : ""
      };
    }

    if (row[7] || row[8]) { 
      normalized.push({
        ...currentBase,
        mainProcess: row[7],
        subprocess: row[8],
        vaNva: row[9],
        cycle: row[10],
        transDay: row[11],
        transMonth: row[12],
        cycleTimeMin: row[13],
        totalWorkloadMin: row[14],
        totalWorkloadHrs: row[15],
        // Aliases for compatibility
        timeStandard: row[13],
        transactionCount: row[12],
        totalTime: row[15],
        fte: isFirstRow ? rowFte : ""
      });
    }
  }
  return normalized;
}

function getProcessDirectory(filters) {
  try {
    const normalizedData = getNormalizedData();
    const user = getUserProfile(); 
    const result = [];
    
    normalizedData.forEach(row => {
      if (row.approvalStatus !== 'Approved') return;

      if (user.role !== 'Admin') {
        if (!isAllowedByScope(user.scope.division, row.division)) return;
        if (!isAllowedByScope(user.scope.group, row.group)) return;
        if (!isAllowedByScope(user.scope.department, row.department)) return;
        if (!isAllowedByScope(user.scope.section, row.section)) return;
      }

      if (filters && filters.division && filters.division !== row.division) return;
      if (filters && filters.group && filters.group !== row.group) return;
      if (filters && filters.department && filters.department !== row.department) return;
      if (filters && filters.section && filters.section !== row.section) return;
      
      result.push({
        mainProcess: String(row.mainProcess).trim(),
        subprocess: String(row.subprocess).trim(),
        vaNva: row.vaNva,
        cycle: row.cycle,
        transDay: Number(row.transDay) || 0,
        transMonth: Number(row.transMonth) || 0,
        cycleTimeMin: Number(row.cycleTimeMin) || 0,
        totalWorkloadMin: Number(row.totalWorkloadMin) || 0,
        totalWorkloadHrs: Number(row.totalWorkloadHrs) || 0,
        namePosition: (row.employeeName ? row.employeeName : 'Vacant') + ' / ' + row.position,
        timeStandard: Number(row.cycleTimeMin) || 0,
        transactionCount: Number(row.transMonth) || 0
      });
    });
    
    result.sort((a, b) => a.mainProcess.localeCompare(b.mainProcess));
    return { success: true, data: result };
  } catch (e) { return { success: false, message: e.message }; }
}

function getSavedPositions(filters) {
  try {
    const normalizedData = getNormalizedData();
    const user = getUserProfile(); 
    const positions = {};
    const reportData = { va: 0, nva: 0, bva: 0 }; 
    
    normalizedData.forEach(row => {
      if (user.role !== 'Admin') {
        if (!isAllowedByScope(user.scope.division, row.division)) return;
        if (!isAllowedByScope(user.scope.group, row.group)) return;
        if (!isAllowedByScope(user.scope.department, row.department)) return;
        if (!isAllowedByScope(user.scope.section, row.section)) return;
      }

      if (filters && filters.division && filters.division !== row.division) return;
      if (filters && filters.group && filters.group !== row.group) return;
      if (filters && filters.department && filters.department !== row.department) return;
      if (filters && filters.section && filters.section !== row.section) return;
      
      if (row.approvalStatus === 'Approved') {
        if (row.vaNva === 'VA') reportData.va += Number(row.totalTime) || 0;
        else if (row.vaNva === 'NVA') reportData.nva += Number(row.totalTime) || 0;
        else if (row.vaNva === 'BVA') reportData.bva += Number(row.totalTime) || 0;
      }

      const key = `${row.position}|${row.employeeName}|${row.division}|${row.group}|${row.department}|${row.section}`;
      if (!positions[key]) {
        positions[key] = {
          position: row.position, employeeName: row.employeeName, division: row.division, 
          group: row.group, department: row.department, section: row.section,
          fte: row.fte || 0, approvalStatus: row.approvalStatus, remarks: row.remarks,
          mainProcesses: new Set([row.mainProcess])
        };
      } else {
        positions[key].mainProcesses.add(row.mainProcess);
        if (row.fte) positions[key].fte = row.fte; 
      }
    });
    
    const result = Object.values(positions).map(p => ({ ...p, mpCount: p.mainProcesses.size }));
    const approvedHc = getApprovedHeadcount(filters, user);
    return { success: true, data: result, reports: reportData, approvedHc: approvedHc };
  } catch (e) { return { success: false, message: e.message }; }
}

function getApprovedHeadcount(filters, user) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orgSheet = ss.getSheetByName('OrgChart');
    if (!orgSheet) return 0;

    const data = orgSheet.getDataRange().getValues();
    if (data.length <= 1) return 0;

    const headers = data[0].map(h => String(h || "").trim().toUpperCase());
    
    let hcCol = headers.findIndex(h => h.includes("APPROVED HC") || h.includes("APPROVED"));
    if (hcCol === -1) hcCol = 4; // Default to Column E (0-indexed 4)

    let divCol = headers.findIndex(h => h.includes("DIVISION"));
    if (divCol === -1) divCol = 0;
    let grpCol = headers.findIndex(h => h.includes("GROUP"));
    if (grpCol === -1) grpCol = 1;
    let deptCol = headers.findIndex(h => h.includes("DEPARTMENT"));
    if (deptCol === -1) deptCol = 2;
    let secCol = headers.findIndex(h => h.includes("SECTION"));
    if (secCol === -1) secCol = 3;

    let totalApproved = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const hcVal = Number(row[hcCol]) || 0;
      if (!hcVal) continue;

      const rDiv = divCol !== -1 ? String(row[divCol] || "").trim() : "";
      const rGrp = grpCol !== -1 ? String(row[grpCol] || "").trim() : "";
      const rDept = deptCol !== -1 ? String(row[deptCol] || "").trim() : "";
      const rSec = secCol !== -1 ? String(row[secCol] || "").trim() : "";

      if (user && user.role !== 'Admin') {
        if (rDiv && !isAllowedByScope(user.scope.division, rDiv)) continue;
        if (rGrp && !isAllowedByScope(user.scope.group, rGrp)) continue;
        if (rDept && !isAllowedByScope(user.scope.department, rDept)) continue;
        if (rSec && !isAllowedByScope(user.scope.section, rSec)) continue;
      }

      if (filters) {
        if (filters.division && rDiv && filters.division.toLowerCase() !== rDiv.toLowerCase()) continue;
        if (filters.group && rGrp && filters.group.toLowerCase() !== rGrp.toLowerCase()) continue;
        if (filters.department && rDept && filters.department.toLowerCase() !== rDept.toLowerCase()) continue;
        if (filters.section && rSec && filters.section.toLowerCase() !== rSec.toLowerCase()) continue;
      }

      totalApproved += hcVal;
    }

    return totalApproved;
  } catch (e) {
    return 0;
  }
}

function getPositionData(pos, empName, div, grp, dept, sec) {
  try {
    const normalizedData = getNormalizedData();
    const records = normalizedData.filter(row => 
      row.position === pos && row.employeeName === empName && row.division === div && 
      row.group === grp && row.department === dept && row.section === sec
    );
    return { success: true, data: records };
  } catch (e) { return { success: false, message: e.message }; }
}

function getDepartmentContext(div, grp, dept, excludePos, excludeEmp) {
  try {
    const normalizedData = getNormalizedData();
    const contextMap = {};
    
    normalizedData.forEach(row => {
      if (row.approvalStatus !== 'Approved') return;
      if (row.division !== div || row.group !== grp || row.department !== dept) return;
      if (row.position === excludePos && row.employeeName === excludeEmp) return;
      
      const key = `${row.employeeName || 'Vacant'} - ${row.position}`;
      if (!contextMap[key]) contextMap[key] = { processes: {} };
      
      if (!contextMap[key].processes[row.mainProcess]) {
        contextMap[key].processes[row.mainProcess] = new Set();
      }
      contextMap[key].processes[row.mainProcess].add(row.subprocess);
    });
    
    const result = {};
    for (const posKey in contextMap) {
      result[posKey] = {};
      for (const mp in contextMap[posKey].processes) {
        result[posKey][mp] = Array.from(contextMap[posKey].processes[mp]);
      }
    }
    
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getPendingApprovals() {
  try {
    const user = getUserProfile();
    // If they are Standard, they don't get an approval pipeline at all
    if (user.role === 'Standard') return { success: true, data: [] };

    const normalizedData = getNormalizedData();
    const pendingSet = new Set();
    const pipelineList = [];
    const myEmail = user.email.toLowerCase();

    normalizedData.forEach(row => {
      let canView = false;
      let isPendingForMe = false;
      
      if (user.role === 'Admin') {
        canView = true;
        isPendingForMe = row.approvalStatus.includes('Pending');
      } else {
        // SMART EMAIL CHECK: Check if they are the L1 Target
        if (row.l1Target === myEmail) {
          canView = true; 
          if (row.approvalStatus === 'Pending Level 1') isPendingForMe = true;
        }
        // SMART EMAIL CHECK: Check if they are the L2 Target
        if (row.l2Target === myEmail) {
          canView = true;
          if (row.approvalStatus === 'Pending Level 2') isPendingForMe = true;
        }
      }

      if (canView) {
        const key = `${row.position}|${row.employeeName}|${row.division}|${row.group}|${row.department}|${row.section}`;
        if (!pendingSet.has(key)) {
          pendingSet.add(key);
          pipelineList.push({
            position: row.position, employeeName: row.employeeName, division: row.division, department: row.department, group: row.group, section: row.section,
            fte: row.fte || 0, status: row.approvalStatus, isPendingForMe: isPendingForMe
          });
        }
      }
    });

    pipelineList.sort((a, b) => {
      if (a.isPendingForMe && !b.isPendingForMe) return -1;
      if (!a.isPendingForMe && b.isPendingForMe) return 1;
      return 0;
    });

    return { success: true, data: pipelineList };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateApprovalStatus(payload) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workload');
    const data = sheet.getDataRange().getValues();
    const { pos, empName, div, grp, dept, sec, action, remarks } = payload;
    const user = getUserProfile();
    
    let newStatus = action === "Reject" ? "Rejected" : "Approved"; 
    let currentBase = {};
    let isDetermined = false;

    for (let i = 1; i < data.length; i++) { 
      const row = data[i];
      const rowPos = String(row[0]).trim();
      const rowFte = String(row[16]).trim();
      
      let isFirstRow = false;
      if (i === 1 || rowFte !== "" || (rowPos !== "" && rowPos !== currentBase.position) || (rowPos !== "" && String(data[i-1][0]).trim() === "")) isFirstRow = true;

      if (isFirstRow) {
        currentBase = {
          position: rowPos || currentBase.position, employeeName: row[17] ? String(row[17]).trim() : "",
          division: String(row[3]).trim() || currentBase.division, group: String(row[4]).trim() || currentBase.group,
          department: String(row[5]).trim() || currentBase.department, section: String(row[6]).trim() || currentBase.section,
          status: String(row[29]).trim() // Grab the current approval status of the row
        };
      }
      
      if (currentBase.position === pos && currentBase.employeeName === empName && currentBase.division === div && currentBase.group === grp && currentBase.department === dept && currentBase.section === sec) {
        
        if (action === "Approve" && !isDetermined) {
          const l1Email = String(data[i][32]).trim().toLowerCase();
          const l2Email = String(data[i][33]).trim().toLowerCase(); 
          
          if (user.role === 'Admin') {
            newStatus = "Approved";
          } else if (currentBase.status === 'Pending Level 1') {
            // Smart bypass: if L2 is the same as L1 (or blank), fully approve it!
            newStatus = (l2Email && l2Email !== l1Email) ? "Pending Level 2" : "Approved";
          } else if (currentBase.status === 'Pending Level 2') {
            newStatus = "Approved";
          }
          
          isDetermined = true; // Lock in the logic so it doesn't recalculate on subprocess rows
        }

        sheet.getRange(i + 1, 30).setValue(isFirstRow ? newStatus : ""); 
        sheet.getRange(i + 1, 31).setValue(isFirstRow ? remarks : "");   
      }
    }
    return { success: true, message: `Workload marked as ${newStatus}!` };
  } catch (e) { 
    return { success: false, message: e.message }; 
  }
}

function getOverlapAnalysis(filters) {
  try {
    const normalizedData = getNormalizedData();
    const user = getUserProfile();
    const mainProcessMap = {}; const subProcessMap = {};

    normalizedData.forEach(row => {
      if (row.approvalStatus !== 'Approved') return;
      if (user.role !== 'Admin') {
        if (!isAllowedByScope(user.scope.division, row.division)) return;
        if (!isAllowedByScope(user.scope.group, row.group)) return;
        if (!isAllowedByScope(user.scope.department, row.department)) return;
        if (!isAllowedByScope(user.scope.section, row.section)) return;
      }
      if (filters.division && filters.division !== row.division) return;
      if (filters.group && filters.group !== row.group) return;
      if (filters.department && filters.department !== row.department) return;
      if (filters.section && filters.section !== row.section) return;

      const mpKey = String(row.mainProcess).trim().toUpperCase();
      const spKey = String(row.subprocess).trim().toUpperCase();
      const posIdentifier = `${row.employeeName || "Vacant"} (${row.position}) - ${row.department}`;

      if (mpKey) {
        if (!mainProcessMap[mpKey]) mainProcessMap[mpKey] = { name: row.mainProcess, details: {} };
        if (!mainProcessMap[mpKey].details[posIdentifier]) mainProcessMap[mpKey].details[posIdentifier] = 0;
        mainProcessMap[mpKey].details[posIdentifier] += Number(row.totalTime) || 0;
      }
      if (spKey) {
        if (!subProcessMap[spKey]) subProcessMap[spKey] = { name: row.subprocess, details: {} };
        if (!subProcessMap[spKey].details[posIdentifier]) subProcessMap[spKey].details[posIdentifier] = 0;
        subProcessMap[spKey].details[posIdentifier] += Number(row.totalTime) || 0;
      }
    });

    const overlaps = [];
    for (const key in mainProcessMap) { if (Object.keys(mainProcessMap[key].details).length > 1) overlaps.push({ process: mainProcessMap[key].name, level: 'Main Process', details: mainProcessMap[key].details }); }
    for (const key in subProcessMap) { if (Object.keys(subProcessMap[key].details).length > 1) overlaps.push({ process: subProcessMap[key].name, level: 'Subprocess', details: subProcessMap[key].details }); }
    return { success: true, data: overlaps };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteWorkloadRecords(pos, empName, div, grp, dept, sec) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workload');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const clean = (val) => String(val || "").trim().toLowerCase();

  const targetPos = clean(pos);
  const targetEmp = clean(empName);
  const targetDiv = clean(div);
  const targetGrp = clean(grp);
  const targetDept = clean(dept);
  const targetSec = clean(sec);

  let rowsToDelete = [];
  let currentBase = {};

  for (let i = 1; i < data.length; i++) { 
    const row = data[i];
    const rowPos = clean(row[0]);
    const rowFte = String(row[16]).trim();
    const prevPos = clean(data[i-1][0]);
    
    let isFirstRow = false;
    if (i === 1) {
      isFirstRow = true;
    } else if (rowFte !== "") {
      isFirstRow = true;
    } else if (rowPos !== "" && rowPos !== currentBase.position) {
      isFirstRow = true;
    } else if (rowPos !== "" && prevPos === "") {
      isFirstRow = true;
    }

    if (isFirstRow) {
      currentBase = {
        position: rowPos || currentBase.position,
        employeeName: row[17] !== undefined ? clean(row[17]) : "",
        division: clean(row[3]) || currentBase.division,
        group: clean(row[4]) || currentBase.group,
        department: clean(row[5]) || currentBase.department,
        section: clean(row[6]) || currentBase.section
      };
    }
    
    if (clean(currentBase.position) === targetPos &&
        clean(currentBase.employeeName) === targetEmp &&
        clean(currentBase.division) === targetDiv &&
        clean(currentBase.group) === targetGrp &&
        clean(currentBase.department) === targetDept &&
        clean(currentBase.section) === targetSec) {
      rowsToDelete.push(i + 1); 
    }
  }
  for (let i = rowsToDelete.length - 1; i >= 0; i--) { sheet.deleteRow(rowsToDelete[i]); }
}

function saveWorkload(payload, positionFTE, isEdit, oldPosData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Workload');
    const user = getUserProfile();
    if (!sheet) return { success: false, message: "Error: 'Workload' sheet not found." };
    if (!payload || payload.length === 0) return { success: false, message: "No subprocesses to save." };

    // Clean up existing records for old position key if edited, and clean up for new position key to prevent duplicates
    if (oldPosData) {
      deleteWorkloadRecords(oldPosData.position, oldPosData.employeeName, oldPosData.division, oldPosData.group, oldPosData.department, oldPosData.section);
    }
    const firstRow = payload[0];
    deleteWorkloadRecords(firstRow.position, firstRow.employeeName, firstRow.division, firstRow.group, firstRow.department, firstRow.section);

    let newStatus = 'Pending Level 1'; 
    if (user.role === 'Admin') newStatus = 'Approved'; 
    else if (!user.l1Approver && user.l2Approver) newStatus = 'Pending Level 2'; 
    
    const rowsToInsert = payload.map((row, index) => [
      index === 0 ? row.position : "",
      index === 0 ? row.fwa : "",
      index === 0 ? row.netReportingHours : "",
      index === 0 ? row.division : "",
      index === 0 ? row.group : "",
      index === 0 ? row.department : "",
      index === 0 ? row.section : "",
      row.mainProcess,
      row.subprocess,
      row.vaNva,
      row.cycle,
      row.transDay,
      row.transMonth,
      row.cycleTimeMin,
      row.totalWorkloadMin,
      row.totalWorkloadHrs,
      index === 0 ? positionFTE : "",
      index === 0 ? row.employeeName : "",
      index === 0 ? row.workLocation : "",
      index === 0 ? row.empType : "",
      index === 0 ? row.agencyName : "",
      index === 0 ? row.empStatus : "",
      index === 0 ? row.startDate : "",
      index === 0 ? row.endDate : "",
      index === 0 ? row.kpi1 : "",
      index === 0 ? row.kpi2 : "",
      index === 0 ? row.kpi3 : "",
      index === 0 ? row.kpi4 : "",
      index === 0 ? row.kpi5 : "",
      index === 0 ? newStatus : "",
      "",
      index === 0 ? user.email : "",
      index === 0 ? user.l1Approver : "",
      index === 0 ? user.l2Approver : ""
    ]);
    
    if (rowsToInsert.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    else return { success: false, message: "No subprocesses to save." };
    
    return { success: true, message: isEdit ? `Workload updated and marked as ${newStatus}!` : `Workload submitted and marked as ${newStatus}!` };
  } catch (error) { return { success: false, message: "Server Error: " + error.message }; }
}

function getOptimizationAnalysis(filters) {
  try {
    const normalizedData = getNormalizedData();
    const user = getUserProfile();

    const orgData = normalizedData.filter(row => {
      if (row.approvalStatus !== 'Approved') return false;
      if (user.role !== 'Admin') {
        if (!isAllowedByScope(user.scope.division, row.division)) return false;
        if (!isAllowedByScope(user.scope.group, row.group)) return false;
        if (!isAllowedByScope(user.scope.department, row.department)) return false;
        if (!isAllowedByScope(user.scope.section, row.section)) return false;
      }
      if (filters.division && filters.division !== row.division) return false;
      if (filters.group && filters.group !== row.group) return false;
      if (filters.department && filters.department !== row.department) return false;
      if (filters.section && filters.section !== row.section) return false;
      return true;
    });

    const positionMap = {}; const nvaTargets = [];
    orgData.forEach(row => {
      const isVacant = !row.employeeName || row.employeeName.trim() === "";
      const empId = isVacant ? `VACANT_${row.position}` : row.employeeName;
      const key = `${empId}|${row.position}`;
      if (!positionMap[key]) positionMap[key] = { position: row.position, employeeName: row.employeeName, isVacant: isVacant, fte: Number(row.fte) || 0, processes: new Set(), totalHours: 0 };
      positionMap[key].processes.add(row.mainProcess);
      positionMap[key].totalHours += Number(row.totalTime) || 0;
      if (row.fte) positionMap[key].fte = Number(row.fte);
      if (row.vaNva === 'NVA') nvaTargets.push({ position: row.position, employeeName: row.employeeName || 'Vacant', task: row.subprocess, mainProcess: row.mainProcess, hours: Number(row.totalTime) || 0 });
    });

    const positions = Object.values(positionMap);
    const vacantRoles = positions.filter(p => p.isVacant);
    const filledRoles = positions.filter(p => !p.isVacant);
    filledRoles.forEach(role => { role.availableCapacity = 1.0 - role.fte; });

    const redistributionPlan = vacantRoles.map(vacant => {
      const candidates = filledRoles.sort((a, b) => b.availableCapacity - a.availableCapacity);
      const topCandidate = candidates.length > 0 ? candidates[0] : null;
      return { vacantPosition: vacant.position, workloadFTE: vacant.fte, topCandidate: topCandidate ? topCandidate.employeeName : "No candidates found", candidatePosition: topCandidate ? topCandidate.position : "N/A", candidateCapacity: topCandidate ? topCandidate.availableCapacity : 0 };
    });

    const nvaSummary = {};
    nvaTargets.forEach(t => {
      const key = `${t.employeeName} (${t.position})`;
      if (!nvaSummary[key]) nvaSummary[key] = { totalHours: 0, tasks: [] };
      nvaSummary[key].totalHours += t.hours;
      nvaSummary[key].tasks.push(t.task);
    });
    const efficiencyPlan = Object.keys(nvaSummary).map(key => ({ entity: key, totalNVAHours: nvaSummary[key].totalHours, tasks: nvaSummary[key].tasks })).sort((a, b) => b.totalNVAHours - a.totalNVAHours);

    return { success: true, redistribution: redistributionPlan, efficiency: efficiencyPlan };
  } catch (e) { return { success: false, message: e.message }; }
}
