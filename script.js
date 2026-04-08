const DEPLOYMENT_URL = "https://script.google.com/macros/s/AKfycbzQyH7gxm9GrLWs4tKz4R-gYOvQ5tZTVIbzXB5p7-qHPdbVJq3iSaBxbxefQ_uJm-xkow/exec";

async function gasRequest(action, payload = {}) {
  if (DEPLOYMENT_URL === "YOUR_DEPLOYMENT_URL_HERE") {
    console.warn("Please set your DEPLOYMENT_URL at the top of script.js");
  }

  const response = await fetch(DEPLOYMENT_URL, {
    method: "POST",
    headers: {
      "Content-Type": "text/plain;charset=utf-8",
    },
    body: JSON.stringify({ action, payload }),
    redirect: "follow"
  });

  if (!response.ok) {
    throw new Error(`HTTP Error: ${response.status}`);
  }

  const out = await response.json();
  if (!out.success) {
    throw new Error(out.error || "Unknown Apps Script Error");
  }

  return out.data;
}

window.currentFormMode = "ADD";
window.isGlobalsLocked = {};  // Changed to object to track per-mode lock state
window.isGlobalsLocked.ADD = localStorage.getItem("isGlobalsLocked_ADD") === "true";
window.isGlobalsLocked.EDIT = localStorage.getItem("isGlobalsLocked_EDIT") === "true";
window.isGlobalsLocked.BULK = localStorage.getItem("isGlobalsLocked_BULK") === "true";

document.addEventListener("DOMContentLoaded", function () {
  const mainContent = document.getElementById("mainContent");
  if (mainContent) mainContent.style.display = "block";

  // Data bindings for Form Page
  const form = document.getElementById("studentForm");
  if (form) {
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      if (currentFormMode === "ADD") {
          saveStudent();
      } else if (currentFormMode === "EDIT") {
          updateStudent();
      }
    });

    // Auto-Save Listeners
    const inputs = form.querySelectorAll("input, select, textarea");
    inputs.forEach(input => {
      input.addEventListener("input", saveFormToCache);
      input.addEventListener("change", saveFormToCache);
    });

    // Check if coming from the database view with edit params
    const params = new URLSearchParams(window.location.search);
    const toEdit = params.get('edit');
    if (toEdit) {
       // Switch visually to EDIT tab first
       setFormMode('EDIT');
       
       const loadingScreen = document.getElementById("loadingScreen");
       
       // Check if we have cached data for instant load
       const cachedEditKey = "editCache_" + toEdit;
       const hasCachedData = localStorage.getItem(cachedEditKey) !== null;
       
       if (hasCachedData) {
           // Show loading screen only briefly since we have cached data
           if (loadingScreen) {
               loadingScreen.style.display = "flex";
               document.querySelector(".status-text").textContent = "Loading from cache...";
           }
       } else {
           // Show loading screen normally for fresh data fetch
           if (loadingScreen) {
               loadingScreen.style.display = "flex";
               document.querySelector(".status-text").textContent = "Loading student data...";
           }
       }
       
       editStudent(toEdit).then(() => {
           if (loadingScreen) loadingScreen.style.display = "none";
       });
    } else {
       // Auto-load Form Draft (Default ADD)
       const loadingScreen = document.getElementById("loadingScreen");
       if (loadingScreen) loadingScreen.style.display = "none";
       loadFormFromCache();
    }
    
    // Apply global lock state
    applyGlobalsLockUI();

    // Bind bulk file input
    const bulkInput = document.getElementById("bulkFileInput");
    if(bulkInput) {
        bulkInput.addEventListener("change", handleBulkFileUpload);
    }
    
    // Bind API Key persistence
    const apiKeyInput = document.getElementById("geminiApiKey");
    if(apiKeyInput) {
        apiKeyInput.value = localStorage.getItem("geminiApiKey") || "";
        apiKeyInput.addEventListener("input", function(e) {
            localStorage.setItem("geminiApiKey", e.target.value.trim());
        });
    }
  }

  // Database Caching and View Logic
  if (document.getElementById("studentTableBody")) {
    loadStudentsWithConnection();
  } else {
    // Hide loading screen if not pulling data immediately
    const loadingScreen = document.getElementById("loadingScreen");
    if (loadingScreen) loadingScreen.style.display = "none";
  }
});

/* =========================================================
   FORM PERSISTENCE & MODE MANAGEMENT
========================================================= */

window.toggleGlobalsLock = function() {
    isGlobalsLocked[currentFormMode] = !isGlobalsLocked[currentFormMode];
    localStorage.setItem("isGlobalsLocked_" + currentFormMode, isGlobalsLocked[currentFormMode]);
    applyGlobalsLockUI();
}

window.lockAdministrativeFields = function() {
    const adminIds = ["college", "ayStart", "ayEnd", "aySemester"];
    adminIds.forEach(id => {
        const el = document.getElementById(id);
        if(el) {
            el.readOnly = true;
            el.classList.add("opacity-50", "cursor-not-allowed", "bg-surface-dim");
        }
    });
}

window.lockSignatoriesFields = function() {
    const signatoryIds = ["preparedBy", "notedBy", "approvedBy"];
    signatoryIds.forEach(id => {
        const el = document.getElementById(id);
        if(el) {
            el.readOnly = true;
            el.classList.add("opacity-50", "cursor-not-allowed");
        }
    });
}

window.unlockAdministrativeFields = function() {
    const adminIds = ["college", "ayStart", "ayEnd", "aySemester"];
    adminIds.forEach(id => {
        const el = document.getElementById(id);
        if(el) {
            el.readOnly = false;
            el.classList.remove("opacity-50", "cursor-not-allowed", "bg-surface-dim");
        }
    });
}

window.unlockSignatoriesFields = function() {
    const signatoryIds = ["preparedBy", "notedBy", "approvedBy"];
    signatoryIds.forEach(id => {
        const el = document.getElementById(id);
        if(el) {
            el.readOnly = false;
            el.classList.remove("opacity-50", "cursor-not-allowed");
        }
    });
}

window.applyGlobalsLockUI = function() {
    const globalIds = ["reportPeriod", "preparedBy", "notedBy", "approvedBy"];
    const icon = document.getElementById("iconLock");
    const text = document.getElementById("textLock");
    const btn = document.getElementById("btnToggleLock");
    
    const isLocked = isGlobalsLocked[currentFormMode] || false;
    
    globalIds.forEach(id => {
        const el = document.getElementById(id);
        if(el) {
            el.readOnly = isLocked;
            if(isLocked) {
                el.classList.add("opacity-50", "cursor-not-allowed", "bg-surface-dim");
            } else {
                el.classList.remove("opacity-50", "cursor-not-allowed", "bg-surface-dim");
            }
        }
    });

    if(icon && text && btn) {
        if(isLocked) {
            icon.className = "bi bi-lock-fill text-error";
            text.textContent = "Locked";
            btn.classList.add("bg-error-container", "text-on-error-container", "border-error/20");
            btn.classList.remove("bg-surface-container-highest", "text-on-surface-variant", "border-outline-variant/50");
        } else {
            icon.className = "bi bi-unlock-fill text-[#e0a800]";
            text.textContent = "Unlocked";
            btn.classList.remove("bg-error-container", "text-on-error-container", "border-error/20");
            btn.classList.add("bg-surface-container-highest", "text-on-surface-variant", "border-outline-variant/50");
        }
    }
}

window.setFormMode = function(mode) {
    if(!document.getElementById("studentForm")) return;
    
    // Guardian: Prevent entering Edit mode randomly
    if (mode === "EDIT") {
        const isFromUrl = new URLSearchParams(window.location.search).has('edit');
        
        // Peek at the edit cache to see if we are currently holding a valid student
        const memRaw = localStorage.getItem("failedReportDraft_EDIT");
        let hasMemory = false;
        if (memRaw) {
            try { hasMemory = !!JSON.parse(memRaw).originalSheetName; } catch (e) {}
        }
        
        if (!isFromUrl && !hasMemory) {
            showToast("⚠️ Select a student from the Students Sheets database to edit.", "warning");
            return; // Block the switch
        }
    }
    
    // Save current state before switching ONLY if standard
    if(currentFormMode !== "BULK") {
        const data = getFormData();
        if(Object.keys(data).length > 0) {
            localStorage.setItem("failedReportDraft_" + currentFormMode, JSON.stringify(data));
        }
    }
    
    window.currentFormMode = mode;
    
    // Update Tab UI
    const tabAdd = document.getElementById("tabAddMode");
    const tabEdit = document.getElementById("tabEditMode");
    const tabBulk = document.getElementById("tabBulkMode");
    const btnSave = document.getElementById("btnSave");
    const btnUpdate = document.getElementById("btnUpdate");
    
    const standardSection = document.getElementById("standardFormSection");
    const bulkSection = document.getElementById("bulkImportSection");
    
    if(mode === "BULK") {
        if(standardSection) standardSection.style.display = "none";
        if(bulkSection) bulkSection.style.display = "block";
    } else {
        if(standardSection) standardSection.style.display = "block";
        if(bulkSection) bulkSection.style.display = "none";
    }
    
    if (mode === "ADD") {
        if(tabAdd) tabAdd.className = "flex-1 py-3 text-sm font-bold rounded-lg bg-primary text-white shadow-md transition-all uppercase tracking-widest flex justify-center items-center gap-2";
        if(tabEdit) tabEdit.className = "flex-1 py-3 text-sm font-bold rounded-lg bg-transparent text-on-surface-variant hover:bg-surface-container-high transition-all uppercase tracking-widest flex justify-center items-center gap-2";
        if(tabBulk) tabBulk.className = "flex-1 py-3 text-sm font-bold rounded-lg bg-transparent text-on-surface-variant hover:bg-surface-container-high transition-all uppercase tracking-widest flex justify-center items-center gap-2";
        if(btnSave) btnSave.style.display = "block";
        if(btnUpdate) btnUpdate.style.display = "none";
        
        // Unlock administrative and signatories fields for fresh ADD operations
        unlockAdministrativeFields();
        unlockSignatoriesFields();
    } else if (mode === "EDIT") {
        if(tabEdit) tabEdit.className = "flex-1 py-3 text-sm font-bold rounded-lg bg-[#e0a800] text-white shadow-md transition-all uppercase tracking-widest flex justify-center items-center gap-2";
        if(tabAdd) tabAdd.className = "flex-1 py-3 text-sm font-bold rounded-lg bg-transparent text-on-surface-variant hover:bg-surface-container-high transition-all uppercase tracking-widest flex justify-center items-center gap-2";
        if(tabBulk) tabBulk.className = "flex-1 py-3 text-sm font-bold rounded-lg bg-transparent text-on-surface-variant hover:bg-surface-container-high transition-all uppercase tracking-widest flex justify-center items-center gap-2";
        if(btnSave) btnSave.style.display = "none";
        if(btnUpdate) btnUpdate.style.display = "block";
    } else if (mode === "BULK") {
        if(tabBulk) tabBulk.className = "flex-1 py-3 text-sm font-bold rounded-lg bg-tertiary-fixed-dim text-tertiary shadow-md transition-all uppercase tracking-widest flex justify-center items-center gap-2";
        if(tabAdd) tabAdd.className = "flex-1 py-3 text-sm font-bold rounded-lg bg-transparent text-on-surface-variant hover:bg-surface-container-high transition-all uppercase tracking-widest flex justify-center items-center gap-2";
        if(tabEdit) tabEdit.className = "flex-1 py-3 text-sm font-bold rounded-lg bg-transparent text-on-surface-variant hover:bg-surface-container-high transition-all uppercase tracking-widest flex justify-center items-center gap-2";
        
        // Unlock administrative and signatories fields for BULK imports
        unlockAdministrativeFields();
        unlockSignatoriesFields();
    }
    
    if (mode !== "BULK") {
        // Clear the visual form first to ensure a clean slate for the memory load
        document.getElementById("studentForm").reset();
        document.getElementById("originalSheetName").value = "";
        if(typeof window.updateVisualToggles === "function") window.updateVisualToggles();
        loadFormFromCache();
    }
    
    // Apply lock UI for the current mode
    applyGlobalsLockUI();
}

function saveFormToCache() {
  const data = getFormData();
  if(Object.keys(data).length > 0) {
    localStorage.setItem("failedReportDraft_" + currentFormMode, JSON.stringify(data));
  }
}

function loadFormFromCache() {
  const draft = localStorage.getItem("failedReportDraft_" + currentFormMode);
  if (!draft) return;
  
  try {
    const data = JSON.parse(draft);
    if(data.college && document.getElementById("college")) {
        document.getElementById("college").value = data.college;
    }
    if(data.academicYearSemester) {
        // Example: "2023-2024 1st Semester" or "2023-2024 Summer"
        const parts = data.academicYearSemester.match(/^(\d+)-(\d+)\s+(.*)$/);
        if(parts) {
            if(document.getElementById("ayStart")) document.getElementById("ayStart").value = parts[1];
            if(document.getElementById("ayEnd")) document.getElementById("ayEnd").value = parts[2];
            if(document.getElementById("aySemester")) document.getElementById("aySemester").value = parts[3];
        }
    }
    const ids = ["studentName", "originalSheetName", "reportPeriod", "courseYear", "subjects", "howFailed", "whatHappened", "whenStarted", "whyHappened", "studentAwareness", "parentAcknowledgment", "remedialTeaching", "performanceAssessment", "activitiesExercisesRemoval", "ifFailedReasons", "preparedBy", "notedBy", "approvedBy", "failedPassed", "stoppedWithdrawDropout"];
    
    ids.forEach(id => {
      const el = document.getElementById(id);
      if (el && data[id]) el.value = data[id];
    });

    if(typeof window.updateVisualToggles === "function") {
        window.updateVisualToggles();
    }
  } catch(e) {
    console.warn("Failed to load draft form: ", e);
  }
}

/* =========================================================
   DATA EXTRACTION
========================================================= */

function getFormData() {
  // Security check mapping
  if(!document.getElementById("studentForm")) return {};

  let ayStart = document.getElementById("ayStart") ? document.getElementById("ayStart").value.trim() : "";
  let ayEnd = document.getElementById("ayEnd") ? document.getElementById("ayEnd").value.trim() : "";
  let aySem = document.getElementById("aySemester") ? document.getElementById("aySemester").value.trim() : "";
  
  let formattedSemester = "";
  if (ayStart || ayEnd || aySem) {
    formattedSemester = `${ayStart}-${ayEnd} ${aySem}`.trim();
  }
  
  return {
    college: document.getElementById("college").value.trim(),
    reportPeriod: document.getElementById("reportPeriod").value.trim(),
    academicYearSemester: formattedSemester,
    studentName: document.getElementById("studentName").value.trim(),
    originalSheetName: document.getElementById("originalSheetName") ? document.getElementById("originalSheetName").value.trim() : "",
    courseYear: document.getElementById("courseYear").value.trim(),
    subjects: document.getElementById("subjects").value.trim(),
    howFailed: document.getElementById("howFailed").value.trim(),
    studentAwareness: document.getElementById("studentAwareness").value.trim(),
    failedPassed: document.getElementById("failedPassed").value.trim(),
    whatHappened: document.getElementById("whatHappened").value.trim(),
    parentAcknowledgment: document.getElementById("parentAcknowledgment").value.trim(),
    ifFailedReasons: document.getElementById("ifFailedReasons").value.trim(),
    whenStarted: document.getElementById("whenStarted").value.trim(),
    remedialTeaching: document.getElementById("remedialTeaching").value.trim(),
    whyHappened: document.getElementById("whyHappened").value.trim(),
    performanceAssessment: document.getElementById("performanceAssessment").value.trim(),
    stoppedWithdrawDropout: document.getElementById("stoppedWithdrawDropout").value.trim(),
    activitiesExercisesRemoval: document.getElementById("activitiesExercisesRemoval").value.trim(),
    preparedBy: document.getElementById("preparedBy").value.trim(),
    notedBy: document.getElementById("notedBy").value.trim(),
    approvedBy: document.getElementById("approvedBy").value.trim()
  };
}

/* =========================================================
   ACTIONS (SAVE, EDIT, UPDATE, DELETE)
========================================================= */

async function saveStudent() {
  const data = getFormData();
  const toast = showToast("Saving report...", "loading", false);
  
  try {
    const res = await gasRequest("addStudent", { data });
    updateToast(toast, res.message || "Successfully sent to Google Sheet!", "success");
    
    // Cache the new student data for future edits
    if (res.sheetName) {
      const cacheKey = "editCache_" + res.sheetName;
      localStorage.setItem(cacheKey, JSON.stringify(data));
    }
    
    // Refresh the cached students list
    try {
      const students = await gasRequest("getAllStudents");
      localStorage.setItem("cachedStudents", JSON.stringify(students));
    } catch (e) {
      console.warn("Could not refresh students cache: ", e);
    }
    
    // Clear only specific student field so the form can be used again
    document.getElementById("studentName").value = "";
    document.getElementById("originalSheetName").value = "";
    
    // Auto-save the new state
    saveFormToCache();
    
  } catch (err) {
    updateToast(toast, err.message, "error");
  }
}

async function updateStudent() {
  const originalSheetName = document.getElementById("originalSheetName").value.trim();
  if (!originalSheetName) {
    showToast("Please select a student first before updating.", "warning");
    return;
  }

  const data = getFormData();
  const toast = showToast("Updating report...", "loading", false);
  
  try {
    const res = await gasRequest("updateStudent", { originalSheetName, data });
    updateToast(toast, res.message || "Update saved successfully!", "success");
    
    // Clear the cached edit data since it's been updated
    localStorage.removeItem("editCache_" + originalSheetName);
    
    // If student name changed, also clear the old cache reference
    const newSheetName = data.studentName ? sanitizeSheetName_(data.studentName) : originalSheetName;
    if (newSheetName !== originalSheetName) {
      localStorage.removeItem("editCache_" + newSheetName);
    }
    
    // Cache the updated data with the new sheet name
    if (res.sheetName) {
      const cacheKey = "editCache_" + res.sheetName;
      localStorage.setItem(cacheKey, JSON.stringify(data));
    }
    
    document.getElementById("studentName").value = "";
    document.getElementById("originalSheetName").value = "";
    saveFormToCache();
    if(typeof window.updateVisualToggles === "function") window.updateVisualToggles();
    
    // Also refresh the cached students list
    loadStudents();
    
    // Optional: bounce back to ADD tab after a successful update?
    // User requested: "snap you back". Let's do it after 1 second.
    setTimeout(() => {
        setFormMode("ADD");
    }, 1500);
    
  } catch (err) {
    updateToast(toast, err.message, "error");
  }
}

// Helper function to sanitize sheet names (matches the Google Apps Script version)
function sanitizeSheetName_(name) {
  return String(name)
    .replace(/[\\\/\?\*\[\]\:]/g, "")
    .trim()
    .substring(0, 99);
}

async function removeStudent(sheetName) {
  if (!confirm(`Are you sure you want to delete "${sheetName}"?`)) return;
  const toast = showToast(`Deleting ${sheetName}...`, "loading", false);
  try {
    const res = await gasRequest("deleteStudent", { sheetName });
    updateToast(toast, res.message || "Deleted successfully.", "success");
    loadStudents();
  } catch (err) {
    updateToast(toast, err.message, "error");
  }
}

async function editStudent(sheetName) {
  if (!document.getElementById("studentForm")) {
    window.location.href = "index.html?edit=" + encodeURIComponent(sheetName);
    return;
  }

  try {
    // Try to load cached edit data first for instant load
    let data = null;
    const cachedEditKey = "editCache_" + sheetName;
    const cachedEdit = localStorage.getItem(cachedEditKey);
    
    if (cachedEdit) {
      try {
        data = JSON.parse(cachedEdit);
        // Use cached data immediately
        populateFormWithData(data);
        window.scrollTo({ top: 0, behavior: "smooth" });
        showToast("Student loaded from cache. Switching to EDIT tab.", "info");
        setFormMode("EDIT");
        saveFormToCache();
        
        // Automatically lock globals for EDIT mode to prevent accidental changes
        if(!isGlobalsLocked["EDIT"]) {
            toggleGlobalsLock();
        }
        
        // Lock administrative and signatories fields from previous edit
        lockAdministrativeFields();
        lockSignatoriesFields();
        
        // Fetch fresh data in background to update cache
        fetchAndCacheStudentData(sheetName);
        return;
      } catch (e) {
        console.warn("Failed to parse cached edit data, fetching fresh copy...");
      }
    }

    // If no cache, fetch from Google Sheet
    data = await gasRequest("getStudent", { sheetName });
    
    // Cache the edit data for next time
    localStorage.setItem(cachedEditKey, JSON.stringify(data));
    
    // Populate form
    populateFormWithData(data);
    
    window.scrollTo({ top: 0, behavior: "smooth" });
    showToast("Student loaded for editing. Switching to EDIT tab.", "info");
    
    setFormMode("EDIT"); // visually lock to EDIT mode
    saveFormToCache(); // Force cache save
    
    // Automatically lock globals for EDIT mode to prevent accidental changes
    if(!isGlobalsLocked["EDIT"]) {
        toggleGlobalsLock();
    }
    
    // Lock administrative and signatories fields from previous edit
    lockAdministrativeFields();
    lockSignatoriesFields();

  } catch (err) {
    showToast(err.message, "error");
  }
}

// Helper function to populate form with data (extracted for reuse)
function populateFormWithData(data) {
  document.getElementById("originalSheetName").value = data.sheetName || "";
  
  // Handle specific dropdowns/split fields
  if(document.getElementById("college")) document.getElementById("college").value = data.college || "";
  
  if (data.academicYearSemester) {
      const parts = data.academicYearSemester.match(/^(\d+)-(\d+)\s+(.*)$/);
      if(parts) {
          if(document.getElementById("ayStart")) document.getElementById("ayStart").value = parts[1];
          if(document.getElementById("ayEnd")) document.getElementById("ayEnd").value = parts[2];
          if(document.getElementById("aySemester")) document.getElementById("aySemester").value = parts[3];
      }
  }

  const ids = ["reportPeriod", "studentName", "courseYear", "subjects", "howFailed", "whatHappened", "whenStarted", "whyHappened", "studentAwareness", "parentAcknowledgment", "remedialTeaching", "performanceAssessment", "activitiesExercisesRemoval", "ifFailedReasons", "preparedBy", "notedBy", "approvedBy", "failedPassed", "stoppedWithdrawDropout"];
  
  ids.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = data[id] || "";
  });
}

// Helper function to fetch and cache student data in background
async function fetchAndCacheStudentData(sheetName) {
  try {
    const data = await gasRequest("getStudent", { sheetName });
    const cachedEditKey = "editCache_" + sheetName;
    localStorage.setItem(cachedEditKey, JSON.stringify(data));
  } catch (err) {
    console.warn("Failed to update cached student data: ", err);
  }
}

function resetForm() {
  localStorage.removeItem("failedReportDraft_" + currentFormMode);
  document.getElementById("studentForm").reset();
  document.getElementById("originalSheetName").value = "";
  
  // Unlock fields when clearing form
  unlockAdministrativeFields();
  unlockSignatoriesFields();
  
  if(typeof window.updateVisualToggles === "function") window.updateVisualToggles();
  showToast("Form cleared and cache wiped.", "info");
}

/* =========================================================
   DATABASE LOADING (SKELETON & CACHE-FIRST MAPPING)
========================================================= */

async function loadStudentsWithConnection() {
  const loadingScreen = document.getElementById("loadingScreen");
  const statusDot = document.querySelector(".status-dot");
  const statusText = document.querySelector(".status-text");

  if(statusText) statusText.textContent = "Connecting...";

  // Cache First Presentation Strategy
  const cached = localStorage.getItem("cachedStudents");
  if (cached && document.getElementById("studentTableBody")) {
     try {
       const students = JSON.parse(cached);
       populateStudentTable(students);
       if(loadingScreen) loadingScreen.style.display = "none";
     } catch (e) {}
  } else if (document.getElementById("studentTableBody")) {
     renderTableSkeleton();
  }

  try {
    const students = await gasRequest("getAllStudents");
    if(statusDot) statusDot.classList.add("connected");
    if(statusText) statusText.textContent = "Connected!";
    
    // Save to Cache
    localStorage.setItem("cachedStudents", JSON.stringify(students));

    setTimeout(() => {
      if (loadingScreen) loadingScreen.style.display = "none";
    }, 1000);

    // Overwrite presentation with fresh data silently 
    if (document.getElementById("studentTableBody")) {
      populateStudentTable(students);
    }
  } catch (err) {
    if(statusDot) statusDot.style.backgroundColor = "#ef4444";
    if(statusText) statusText.textContent = "Connection failed. Retrying...";
    console.error(err);
    setTimeout(loadStudentsWithConnection, 2000);
  }
}

async function loadStudents() {
  if (!document.getElementById("studentTableBody")) return;
  
  // Re-render skeletons visually for manual refresh
  renderTableSkeleton();
  
  try {
    const students = await gasRequest("getAllStudents");
    localStorage.setItem("cachedStudents", JSON.stringify(students));
    populateStudentTable(students);
  } catch (err) {
    showToast(err.message, "error");
  }
}

/* =========================================================
   BULK IMPORT QUEUE
========================================================= */

window.bulkStudents = [];
window.isBulkUploading = false;

window.processSmartText = function() {
    const rawText = document.getElementById("bulkPasteInput").value;
    if(!rawText || rawText.trim() === "") {
        showToast("Please paste some text into the box first.", "warning");
        return;
    }
    
    // Normalize newlines
    let text = rawText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    
    // Split text by student landmarks (use courseYear as marker for new student blocks)
    // This ensures we don't start mid-table where student name might be inside a table
    const splitRegex = /(?=COURSE & YEAR:|Course & Year:|ACADEMIC YEAR|Academic Year)/i;
    let studentBlocks = text.split(splitRegex);
    
    // Filter empty blocks - keep only blocks that have substantial content
    studentBlocks = studentBlocks.filter(block => block.match(/NAME OF STUDENT:|Name of the Student:|SUBJECT/i));
    
    if(studentBlocks.length === 0) {
        showToast("Could not find any student markers. Ensure it contains 'NAME OF STUDENT:' or 'COURSE & YEAR:'.", "error");
        return;
    }
    
    window.bulkStudents = [];
    
    studentBlocks.forEach(block => {
        const markers = [
            { key: "studentName", terms: ["NAME OF STUDENT:", "Name of the Student:"] },
            { key: "courseYear", terms: ["Course & Year:"] },
            { key: "subjects", terms: ["SUBJECT/S"] },
            { key: "howFailed", terms: ["HOW DID HE/SHE FAIL?", "HOW DID HE/SHE FAIL"] },
            { key: "whatHappened", terms: ["WHAT HAPPENED?", "WHAT HAPPENED"] },
            { key: "whenStarted", terms: ["WHEN IT WAS STARTED?", "WHEN IT WAS STARTED"] },
            { key: "whyHappened", terms: ["WHY DID IT HAPPEN?", "WHY DID IT HAPPEN"] },
            { key: "stoppedWithdrawDropout", terms: ["STOPPED/WITHDRAW/DROP OUT", "STOPPED/WITHDRAW/DROPOUT"] },
            { key: "studentAwareness", terms: ["STUDENT’S AWARENESS", "STUDENT'S AWARENESS", "STUDENTS AWARENESS"] },
            { key: "parentAcknowledgment", terms: ["PARENT’S ACKNOWLEDGMENT", "PARENT'S ACKNOWLEDGMENT", "PARENTS ACKNOWLEDGMENT"] },
            { key: "remedialTeaching", terms: ["REMEDIAL TEACHING"] },
            { key: "performanceAssessment", terms: ["PERFORMANCE ASSESSMENT"] },
            { key: "activitiesExercisesRemoval", terms: ["ANY GIVEN ACTIVITIES, EXERCISES OR REMOVAL EXAMS?", "ANY GIVEN ACTIVITIES"] },
            { key: "ifFailedReasons", terms: ["IF FAILED, REASONS", "IF FAILED REASONS"] },
            { key: "failedPassed", terms: ["FAILED/PASSED"] }
        ];
        
        let extracted = {};
        let located = [];
        
        // Find indices of markers within the text block
        markers.forEach(m => {
            let earliestIdx = -1;
            let matchLen = 0;
            m.terms.forEach(t => {
                const idx = block.toUpperCase().indexOf(t.toUpperCase());
                if(idx !== -1 && (earliestIdx === -1 || idx < earliestIdx)) {
                    earliestIdx = idx;
                    matchLen = t.length;
                }
            });
            if(earliestIdx !== -1) {
                located.push({ key: m.key, idx: earliestIdx, len: matchLen });
            }
        });
        
        // Sort sequentially
        located.sort((a,b) => a.idx - b.idx);
        
        // Extract substring between current marker and next marker
        for(let i=0; i < located.length; i++) {
            const startStr = located[i].idx + located[i].len;
            let endStr = block.length;
            if(i+1 < located.length) {
                endStr = located[i+1].idx;
            }
            
            let rawVal = block.substring(startStr, endStr);
            // aggressively strip underscores/blank lines
            rawVal = rawVal.replace(/^[_\s]+/g, '').replace(/[_\s]+$/g, '').trim(); 
            extracted[located[i].key] = rawVal;
        }

        // Inherit global defaults from current DOM State (ADMINISTRATIVE & SIGNATORIES)
        extracted.reportPeriod = document.getElementById("reportPeriod")?.value || "";
        extracted.preparedBy = document.getElementById("preparedBy")?.value || "";
        extracted.notedBy = document.getElementById("notedBy")?.value || "";
        extracted.approvedBy = document.getElementById("approvedBy")?.value || "";
        extracted.college = document.getElementById("college")?.value || "";
        
        const p1 = document.getElementById("ayStart")?.value.trim() || "";
        const p2 = document.getElementById("ayEnd")?.value.trim() || "";
        const p3 = document.getElementById("aySemester")?.value.trim() || "";
        let formattedSemester = "";
        if(p1 && p2 && p3) {
            formattedSemester = `${p1}-${p2} ${p3}`;
        }
        extracted.academicYearSemester = formattedSemester;
        
        if(!extracted.failedPassed) { extracted.failedPassed = "FAILED"; }
        
        // Final sanity check
        if(extracted.studentName && extracted.studentName.length > 0) {
            window.bulkStudents.push({ data: extracted, status: "pending" });
        }
    });
    
    if(window.bulkStudents.length === 0) {
        showToast("Extracted 0 valid students. Check text formatting.", "error");
        return;
    }
    
    document.getElementById("bulkPreviewContainer").style.display = "block";
    document.getElementById("bulkCount").textContent = window.bulkStudents.length;
    renderBulkPreview();
    showToast(`Successfully extracted ${window.bulkStudents.length} students.`, "success");
}

function renderBulkPreview() {
    const tbody = document.getElementById("bulkPreviewBody");
    if(!tbody) return;
    
    tbody.innerHTML = window.bulkStudents.map((item, index) => {
        let statusBadge = '<span class="px-2 py-1 bg-surface-container-highest text-on-surface-variant text-xs font-bold rounded">Pending</span>';
        if(item.status === "uploading") {
            statusBadge = '<span class="px-2 py-1 bg-[#1a237e]/10 text-[#1a237e] text-xs font-bold rounded animate-pulse"><i class="bi bi-arrow-repeat animate-spin inline-block mr-1"></i>Sending</span>';
        } else if (item.status === "success") {
            statusBadge = '<span class="px-2 py-1 bg-[#4ade80]/20 text-[#16a34a] text-xs font-bold rounded"><i class="bi bi-check-lg mr-1"></i>Done</span>';
        } else if (item.status === "error") {
            statusBadge = '<span class="px-2 py-1 bg-error/10 text-error text-xs font-bold rounded" title="Failed to upload."><i class="bi bi-x-lg mr-1"></i>Failed</span>';
        }
        
        return `
            <tr class="hover:bg-surface-container-low transition-colors">
                <td class="p-4 text-center text-on-surface-variant font-bold">${index + 1}</td>
                <td class="p-4">${statusBadge}</td>
                <td class="p-4 font-bold text-primary">${escapeHtml(item.data.studentName)}</td>
                <td class="p-4 text-on-surface-variant text-sm">${escapeHtml(item.data.college)}</td>
                <td class="p-4 text-on-surface-variant text-sm">${escapeHtml(item.data.courseYear)}</td>
            </tr>
        `;
    }).join('');
}

window.startBulkUpload = async function() {
    if(window.isBulkUploading) return;
    if(window.bulkStudents.length === 0) return;
    
    window.isBulkUploading = true;
    const btn = document.getElementById("btnStartBulkUpload");
    btn.innerHTML = '<i class="bi bi-arrow-repeat animate-spin mr-2 inline-block"></i>Queue Running...';
    btn.classList.add("opacity-50", "cursor-not-allowed");
    
    for(let i = 0; i < window.bulkStudents.length; i++) {
        const item = window.bulkStudents[i];
        if(item.status === "success") continue;
        
        item.status = "uploading";
        renderBulkPreview();
        
        try {
            await gasRequest("addStudent", { data: item.data });
            item.status = "success";
        } catch(err) {
            item.status = "error";
            console.error(err);
        }
        renderBulkPreview();
    }
    
    window.isBulkUploading = false;
    btn.innerHTML = '<i class="bi bi-check-all mr-2"></i>Queue Finished';
    showToast("Bulk upload queue has completed.", "info");
}

function renderTableSkeleton() {
  const tbody = document.getElementById("studentTableBody");
  if(!tbody) return;
  
  let rows = "";
  for(let i=0; i<3; i++) {
    rows += `
      <tr class="animate-pulse border-b border-surface-container-high/50">
        <td class="p-5"><div class="h-4 bg-surface-container-highest rounded w-3/4"></div></td>
        <td class="p-5"><div class="h-4 bg-surface-container-highest rounded w-1/2"></div></td>
        <td class="p-5"><div class="h-4 bg-surface-container-highest rounded w-2/3"></div></td>
        <td class="p-5"><div class="h-8 bg-surface-container-highest rounded-md w-20"></div></td>
        <td class="p-5">
            <div class="flex gap-2">
                <div class="h-8 bg-surface-container-highest rounded-md w-14"></div>
                <div class="h-8 bg-surface-container-highest rounded-md w-14"></div>
            </div>
        </td>
      </tr>
    `;
  }
  tbody.innerHTML = rows;
}

function populateStudentTable(students) {
  const tbody = document.getElementById("studentTableBody");
  if (!tbody) return;

  if (!students || students.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="5" class="p-8 text-center text-on-surface-variant font-medium">No student sheets found.</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = students.map(st => `
    <tr class="hover:bg-surface-container-low transition-colors border-b border-surface-container-high/50 last:border-0">
      <td class="p-4 font-bold text-primary">${escapeHtml(st.studentName || st.sheetName)}</td>
      <td class="p-4 text-sm text-on-surface-variant">${escapeHtml(st.courseYear || "")}</td>
      <td class="p-4 text-sm text-on-surface-variant">${escapeHtml(st.college || "")}</td>
      <td class="p-4"><span class="px-3 py-1 bg-surface-container-highest text-on-surface-variant text-xs font-bold rounded-lg shadow-sm border border-outline-variant/30">${escapeHtml(st.status || "N/A")}</span></td>
      <td class="p-4">
        <div class="flex gap-2 flex-wrap">
          <button type="button" class="px-4 py-2 bg-[#e0a800] text-white text-xs font-bold rounded-lg shadow-sm hover:brightness-110 transition-all" onclick="editStudent('${escapeAttr(st.sheetName)}')">Edit</button>
          <button type="button" class="px-4 py-2 bg-error text-white text-xs font-bold rounded-lg shadow-sm hover:bg-error-container hover:text-error transition-all" onclick="removeStudent('${escapeAttr(st.sheetName)}')">Delete</button>
        </div>
      </td>
    </tr>
  `).join("");
}

/* =========================================================
   TOAST NOTIFICATION ENGINE
========================================================= */

window.showMessage = function(msg, type) {
    showToast(msg, type);
};

function showToast(message, type = "info", autoRemove = true) {
  const container = document.getElementById("toastContainer");
  if(!container) return null;
  
  const toast = document.createElement("div");
  
  let bgColor = "bg-surface-container-highest text-on-surface border border-outline-variant/30";
  let icon = '<i class="bi bi-info-circle-fill text-xl text-primary"></i>';
  
  if (type === "success") {
      bgColor = "bg-[#e8f5e9] border border-[#a5d6a7] text-[#2e7d32]";
      icon = '<i class="bi bi-check-circle-fill text-xl text-[#2e7d32]"></i>';
  } else if (type === "error") {
      bgColor = "bg-error-container border border-error/50 text-on-error-container";
      icon = '<i class="bi bi-x-circle-fill text-xl text-error"></i>';
  } else if (type === "warning") {
      bgColor = "bg-[#fff8e1] border border-[#ffe082] text-[#f57f17]";
      icon = '<i class="bi bi-exclamation-triangle-fill text-xl text-[#f57f17]"></i>';
  } else if (type === "loading") {
      bgColor = "bg-primary-fixed border border-primary/20 text-on-primary-fixed-variant";
      icon = '<i class="bi bi-arrow-repeat text-xl animate-spin text-primary"></i>';
  }

  toast.className = `flex items-center gap-3 px-5 py-4 rounded-xl shadow-lg pointer-events-auto transition-all duration-300 transform translate-x-full ${bgColor}`;
  toast.innerHTML = `
    <div class="flex-shrink-0">${icon}</div>
    <span class="font-semibold text-sm tracking-wide">${escapeHtml(message)}</span>
  `;
  
  container.appendChild(toast);
  
  // Slide in
  setTimeout(() => {
    toast.classList.remove("translate-x-full");
  }, 10);
  
  if (autoRemove) {
    setTimeout(() => removeToast(toast), 5000);
  }
  
  return toast;
}

function updateToast(toast, message, type, autoRemove = true) {
  if(!toast) return;
  
  let bgColor = "bg-surface-container-highest text-on-surface border border-outline-variant/30";
  let icon = '<i class="bi bi-info-circle-fill text-xl text-primary"></i>';
  
  if (type === "success") {
      bgColor = "bg-[#e8f5e9] border border-[#a5d6a7] text-[#2e7d32]";
      icon = '<i class="bi bi-check-circle-fill text-xl text-[#2e7d32]"></i>';
  } else if (type === "error") {
      bgColor = "bg-error-container border border-error/50 text-on-error-container";
      icon = '<i class="bi bi-x-circle-fill text-xl text-error"></i>';
  }
  
  toast.className = `flex items-center gap-3 px-5 py-4 rounded-xl shadow-lg pointer-events-auto transition-all duration-300 ${bgColor}`;
  toast.innerHTML = `
    <div class="flex-shrink-0">${icon}</div>
    <span class="font-semibold text-sm tracking-wide">${escapeHtml(message)}</span>
  `;
  
  if (autoRemove) {
    setTimeout(() => removeToast(toast), 5000);
  }
}

function removeToast(toast) {
  if(!toast) return;
  toast.classList.add("translate-x-full", "opacity-0");
  setTimeout(() => {
    if(toast.parentElement) toast.remove();
  }, 300);
}

/* =========================================================
   UTILITIES
========================================================= */

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttr(str) {
  return String(str || "")
    .replace(/'/g, "\\'");
}

function copyConversionPrompt() {
  const prompt = `PROMPT FOR CONVERTING FAILED STUDENTS' REPORT TO SMART TEXT READER FORMAT

Convert the Failed Students' Report from table format to Smart Text Reader format. Follow these rules:

TRANSFORMATION RULES:

1. Extract data from the table structure and convert to header-based format with markers (colon after each field name)

2. Required Headers (in this exact order):
   - NAME OF STUDENT:
   - COURSE & YEAR:
   - SUBJECT/S:
   - HOW DID HE/SHE FAIL?:
   - WHAT HAPPENED?:
   - WHEN IT WAS STARTED?:
   - WHY DID IT HAPPEN?:
   - STOPPED/WITHDRAW/DROP OUT:
   - STUDENT'S AWARENESS:
   - PARENT'S ACKNOWLEDGMENT:
   - REMEDIAL TEACHING:
   - PERFORMANCE ASSESSMENT:
   - ANY GIVEN ACTIVITIES, EXERCISES OR REMOVAL EXAMS?:
   - IF FAILED, REASONS:
   - FAILED/PASSED

3. COURSE & YEAR format: Always "BS Information Technology - 1st Year"

4. Narration Integration: Merge standalone narration explanations INTO the "IF FAILED, REASONS:" field as the closing explanation (do NOT include a "NARRATION:" label)

5. STUDENT'S AWARENESS personalization:
   - For most students: "I informed the [SURNAME] to come to my office, but no one showed up." (use student's surname from NAME OF STUDENT field)
   - For transfer students (Bajao, Gamboa): "[SURNAME] informed classmates about the transfer reasons."
   - For students who were notified (Vidal): "Instructor notified [SURNAME] about the failing grades."
   - For unreachable students (Bautista, Manatad): "[SURNAME] was completely unreachable with no explanation for discontinuing."

6. WHEN IT WAS STARTED?: Extract the date exactly as provided

7. Consolidate multi-line explanations into single clear sentences

8. Remove all table formatting, administrative headers, and signature blocks - keep only student data

9. No separator lines between student records`;

  navigator.clipboard.writeText(prompt).then(() => {
    showToast("✓ Conversion prompt copied to clipboard!", "success", true);
  }).catch(err => {
    console.error("Failed to copy:", err);
    showToast("✗ Failed to copy prompt. Try again.", "error", true);
  });
}
