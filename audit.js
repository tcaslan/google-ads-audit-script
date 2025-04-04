// Google Ads Audit Script
// Version: 1.9 (Added Author Info, removed conditional formatting due to API errors)
// Date: 2025-04-04
// Author: Tugaycan Aslan (https://www.linkedin.com/in/tugaycanaslan/ | https://x.com/tugaycanaslan)
// Generated & Modified with Cline
// Description: Performs a comprehensive, read-only audit of a Google Ads account
//              based on a predefined checklist. Logs results and exports a dashboard-style
//              summary to a Google Sheet with multiple tabs. Handles API limitations gracefully.

// --- Configuration ---
// Customize these thresholds according to your account goals and standards.
var CONFIG = {
  // Performance Thresholds
  MIN_QUALITY_SCORE: 5, // Minimum acceptable Quality Score
  MAX_CPA: 50.0, // Maximum acceptable Cost Per Acquisition ($) - Used for Account level check only now
  MIN_CTR: 0.01, // Minimum acceptable Click-Through Rate (1%) - Used for Account level check only now
  MIN_CONVERSION_RATE: 0.01, // Minimum acceptable Conversion Rate (1%) - Used for Account level check only now
  MIN_IMPRESSION_SHARE: 0.6, // Minimum acceptable Impression Share (60%)
  MAX_IMPRESSION_SHARE_LOST_RANK: 0.2, // Maximum acceptable IS Lost (Rank) (20%)
  MAX_IMPRESSION_SHARE_LOST_BUDGET: 0.1, // Maximum acceptable IS Lost (Budget) (10%)

  // Structure Thresholds
  MIN_ADS_PER_ADGROUP: 2, // Minimum number of active ads per ad group
  MAX_KEYWORDS_PER_ADGROUP: 20, // Maximum recommended keywords per ad group (guideline)

  // Naming Conventions (Examples - use regex patterns)
  CAMPAIGN_NAMING_CONVENTION_REGEX: /^[A-Z]{2,}-[A-Za-z0-9]+-.+$/, // e.g., US-Brand-Search
  ADGROUP_NAMING_CONVENTION_REGEX: /^[A-Za-z0-9]+_.+$/, // e.g., General_Keywords

  // Reporting
  SPREADSHEET_NAME_PREFIX: "Google_Ads_Audit_", // Prefix for the output Google Sheet
  DATE_FORMAT: "yyyy-MM-dd", // Date format for the spreadsheet name

  // Other
  CHECK_LANDING_PAGES: true, // Set to false to skip landing page checks (can be time-consuming)
  LANDING_PAGE_SAMPLE_SIZE: 100, // Number of landing pages to check per campaign (if CHECK_LANDING_PAGES is true)
};

// --- Global Variables ---
var SPREADSHEET_URL = null;
var SPREADSHEET_ID = null; // Store spreadsheet ID for easier access
var ALL_RESULTS = {}; // Object to hold results categorized by sheet name { sheetName: [ [row], [row], ... ] }
var CRITICAL_ISSUES = []; // Array to hold 'Fail' status items for the Overview sheet

// Define Sheet Names
var SHEET_NAMES = {
  OVERVIEW: "Overview",
  PERFORMANCE: "Performance Summary",
  STRUCTURE_SETTINGS: "Structure & Settings",
  KEYWORDS_ADGROUPS: "Keywords & AdGroups",
  ADS_EXTENSIONS: "Ads & Extensions",
  MANUAL_CHECKS: "Opportunities & Manual Checks"
};


/**
 * Main function to orchestrate the Google Ads account audit.
 */
function main() {
  Logger.log("Starting Google Ads Account Audit...");

  // Initialize the Google Sheet for reporting
  initializeSpreadsheet();
  if (!SPREADSHEET_ID) { // Check SPREADSHEET_ID now
    Logger.log("Failed to initialize spreadsheet. Aborting audit.");
    return;
  }
  Logger.log("Audit results will be saved to: " + SPREADSHEET_URL);

  // Run audit modules for each category
  auditAccountSettings();
  auditAccountStructure();
  auditConversionTracking();
  auditKeywords();
  auditAdGroups();
  auditAdCopy();
  auditAdExtensions();
  auditBiddingStrategies();
  auditQualityScore();
  if (CONFIG.CHECK_LANDING_PAGES) {
    auditLandingPages();
  } else {
    addResult("Landing Pages", "Landing Page Checks", "Skipped", "CONFIG.CHECK_LANDING_PAGES is false", "Enable in CONFIG if needed.");
  }
  auditAudienceTargeting();
  auditPerformanceMetrics();
  auditCampaignOptimization();
  auditAutomationTools();
  auditCompetitiveAnalysis();
  auditReportingInsights(); // Mostly notes manual checks

  // Populate the Overview sheet with summaries
  populateOverviewSheet();

  // Write collected results to the respective sheets
  writeResultsToSpreadsheet();

  Logger.log("Google Ads Account Audit Completed.");
  Logger.log("Audit summary saved to: " + SPREADSHEET_URL);
}

// --- Spreadsheet Functions ---

/**
 * Creates a new Google Sheet with predefined tabs or opens an existing one.
 * Clears old data from sheets before populating.
 */
function initializeSpreadsheet() {
    try {
        if (typeof DriveApp === 'undefined' || typeof SpreadsheetApp === 'undefined') {
            Logger.log("Error: DriveApp or SpreadsheetApp is not available.");
            throw new Error("Required Google Apps Script services not available.");
        }

        var dateStr = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), CONFIG.DATE_FORMAT);
        var spreadsheetName = CONFIG.SPREADSHEET_NAME_PREFIX + dateStr;
        var spreadsheet;

        var files = DriveApp.getFilesByName(spreadsheetName);
        if (files.hasNext()) {
            spreadsheet = SpreadsheetApp.open(files.next());
            SPREADSHEET_ID = spreadsheet.getId();
            Logger.log("Using existing spreadsheet: " + spreadsheetName);
        } else {
            spreadsheet = SpreadsheetApp.create(spreadsheetName);
            SPREADSHEET_ID = spreadsheet.getId();
            Logger.log("Created new spreadsheet: " + spreadsheetName);
            // Rename the default "Sheet1" to Overview
            try { spreadsheet.getSheetByName('Sheet1').setName(SHEET_NAMES.OVERVIEW); } catch(e) {/* Ignore if already renamed */}
        }

        SPREADSHEET_URL = spreadsheet.getUrl();

        // Ensure all required sheets exist, clear old data, and set headers
        var requiredSheets = Object.values(SHEET_NAMES);
        var existingSheets = spreadsheet.getSheets().map(function(s) { return s.getName(); });

        requiredSheets.forEach(function(sheetName) {
            var sheet;
            if (existingSheets.indexOf(sheetName) === -1) {
                sheet = spreadsheet.insertSheet(sheetName);
                Logger.log("Created sheet: " + sheetName);
            } else {
                sheet = spreadsheet.getSheetByName(sheetName);
            }

            // Clear content below header row (Row 1)
            if (sheet.getLastRow() > 1) {
                sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
            }
            // Clear formatting below header row
             if (sheet.getLastRow() > 1) {
                 sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearFormat();
             }


            // Set headers if sheet is completely empty or only header exists and is empty
            if (sheet.getLastRow() === 0 || (sheet.getLastRow() === 1 && sheet.getRange("A1").getValue() === "")) {
                 if (sheet.getLastRow() === 1) sheet.clearContents(); // Clear empty header if present
                 var headers;
                 if (sheetName === SHEET_NAMES.OVERVIEW) {
                     headers = ["Category / Item", "Status / Details", "Recommendation"]; // Headers for Overview
                 } else {
                     headers = ["Category", "ChecklistItem", "Status", "Details/Metrics", "Recommendation"]; // Headers for detail sheets
                 }
                 sheet.appendRow(headers);
                 sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
                 sheet.setFrozenRows(1);
            }
            // Initialize results array for this sheet
            ALL_RESULTS[sheetName] = [];
        });

        // Note: Attempting to reorder sheets (e.g., making Overview first)
        // using getSheetIndex() and moveActiveSheet() can cause errors in Ads Scripts.
        // The sheets will be created, but their order might vary.

    } catch (e) {
        Logger.log("Error initializing spreadsheet: " + e);
        if (e.message) {
            Logger.log("Error message: " + e.message);
        }
        if (e.stack) {
            Logger.log("Error stack: " + e.stack);
        }
        SPREADSHEET_URL = null;
        SPREADSHEET_ID = null;
    }
}

/**
 * Determines the target sheet name based on the audit category.
 * @param {string} category - The audit category.
 * @return {string} The name of the sheet to add the result to.
 */
function getSheetNameForCategory(category) {
    switch (category) {
        case "Account Settings":
        case "Account Structure":
            return SHEET_NAMES.STRUCTURE_SETTINGS;
        case "Keywords":
        case "Ad Groups":
        case "Quality Score":
            return SHEET_NAMES.KEYWORDS_ADGROUPS;
        case "Ad Copy":
        case "Ad Extensions (Assets)":
            return SHEET_NAMES.ADS_EXTENSIONS;
        case "Performance Metrics":
        case "Bidding Strategies":
        case "Campaign Optimization":
            return SHEET_NAMES.PERFORMANCE;
        case "Conversion Tracking": // Often manual checks now
        case "Landing Pages": // Includes manual checks
        case "Audience Targeting": // Includes manual checks
        case "Automation & Tools":
        case "Competitive Analysis":
        case "Reporting & Insights":
            return SHEET_NAMES.MANUAL_CHECKS;
        default:
            Logger.log("Warning: Unknown category '" + category + "'. Placing in Manual Checks sheet.");
            return SHEET_NAMES.MANUAL_CHECKS; // Default fallback
    }
}

/**
 * Adds a result row to the appropriate category array and logs it.
 * Also adds critical issues to a separate list.
 * @param {string} category - The audit category (e.g., "Keywords").
 * @param {string} item - The specific checklist item (e.g., "Negative Keywords").
 * @param {string} status - "Pass", "Fail", "Warn", "Info", "Error", "Skipped".
 * @param {string} details - Specific metrics or findings.
 * @param {string} recommendation - Actionable advice.
 */
function addResult(category, item, status, details, recommendation) {
    var rowData = [category, item, status, details, recommendation];
    var sheetName = getSheetNameForCategory(category);

    // Ensure the array for the sheet exists
    if (!ALL_RESULTS[sheetName]) {
        ALL_RESULTS[sheetName] = [];
    }
    ALL_RESULTS[sheetName].push(rowData);

    // Add to critical issues list if status is Fail
    if (status === "Fail") {
        CRITICAL_ISSUES.push(rowData);
    }

    // Log immediately as well
    Logger.log("[" + status + "] " + category + " - " + item + ": " + details + " (" + recommendation + ")");
}

/**
 * Populates the Overview sheet with summary data.
 */
function populateOverviewSheet() {
    if (!SPREADSHEET_ID) return;
    Logger.log("Populating Overview Sheet...");

    try {
        var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        var overviewSheet = spreadsheet.getSheetByName(SHEET_NAMES.OVERVIEW);
        // Clear content below header row (Row 1)
        if (overviewSheet.getLastRow() > 1) {
            overviewSheet.getRange(2, 1, overviewSheet.getLastRow() - 1, overviewSheet.getMaxColumns()).clearContent().clearFormat();
        }
        overviewSheet.setColumnWidths(1, 3, 400); // Reset column widths after clearing

        var failCount = 0;
        var warnCount = 0;

        // --- Summary Section ---
        overviewSheet.appendRow(["--- AUDIT SUMMARY ---", "", ""]);
        var summaryHeaderRange = overviewSheet.getRange(overviewSheet.getLastRow(), 1, 1, 3);
        summaryHeaderRange.merge().setFontWeight("bold").setBackground("#D3D3D3").setHorizontalAlignment("center");


        // Calculate Fails/Warns per category
        var issueCounts = {};
        var sheetOrder = [ // Define order for summary
            SHEET_NAMES.PERFORMANCE,
            SHEET_NAMES.STRUCTURE_SETTINGS,
            SHEET_NAMES.KEYWORDS_ADGROUPS,
            SHEET_NAMES.ADS_EXTENSIONS,
            SHEET_NAMES.MANUAL_CHECKS
        ];

        sheetOrder.forEach(function(sheetName) {
            if (!ALL_RESULTS[sheetName]) return; // Skip if no results for this sheet

            var categoryFails = 0;
            var categoryWarns = 0;
            ALL_RESULTS[sheetName].forEach(function(row) {
                if (row[2] === "Fail") categoryFails++;
                if (row[2] === "Warn") categoryWarns++;
            });
            failCount += categoryFails;
            warnCount += categoryWarns;
            if (categoryFails > 0 || categoryWarns > 0) {
                 issueCounts[sheetName] = { fails: categoryFails, warns: categoryWarns };
            }
        });

        var summaryStartRow = overviewSheet.getLastRow() + 1;
        overviewSheet.appendRow(["Total Critical Issues (Fail)", failCount, "Review items marked 'Fail' below and in respective sheets."]);
        overviewSheet.appendRow(["Total Warnings (Warn)", warnCount, "Review items marked 'Warn' in respective sheets."]);
        overviewSheet.appendRow(["Issue Counts by Area:", "", ""]); // Sub-header

        sheetOrder.forEach(function(sheetName) {
             if (issueCounts[sheetName]) {
                 overviewSheet.appendRow([sheetName, "Fails: " + issueCounts[sheetName].fails + ", Warns: " + issueCounts[sheetName].warns, "See '" + sheetName + "' sheet for details."]);
             }
        });
         overviewSheet.appendRow(["", "", ""]); // Spacer


        // --- Critical Issues Section ---
         overviewSheet.appendRow(["--- CRITICAL ISSUES (FAIL) ---", "", ""]);
         var criticalHeaderRange = overviewSheet.getRange(overviewSheet.getLastRow(), 1, 1, 3);
         criticalHeaderRange.merge().setFontWeight("bold").setBackground("#FFCCCB").setHorizontalAlignment("center"); // Light red background

        if (CRITICAL_ISSUES.length > 0) {
            CRITICAL_ISSUES.forEach(function(issue) {
                // Format: [Category - Item, Details, Recommendation]
                overviewSheet.appendRow([issue[0] + " - " + issue[1], issue[3], issue[4]]);
            });
        } else {
            overviewSheet.appendRow(["No critical 'Fail' issues found.", "", ""]);
        }

        // Apply formatting to summary counts
        overviewSheet.getRange(summaryStartRow, 2).setBackground(failCount > 0 ? "#FFCCCB" : "#90EE90").setFontWeight("bold"); // Red if fails, green if none
        overviewSheet.getRange(summaryStartRow + 1, 2).setBackground(warnCount > 0 ? "#FFFFE0" : "#90EE90").setFontWeight("bold"); // Yellow if warns, green if none


    } catch (e) {
        Logger.log("Error populating overview sheet: " + e);
    }

}


/**
 * Writes all collected audit results to the respective sheets in the Google Sheet.
 */
function writeResultsToSpreadsheet() {
    if (!SPREADSHEET_ID) {
        Logger.log("Spreadsheet not initialized. Cannot write results.");
        return;
    }

    try {
        var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

        for (var sheetName in ALL_RESULTS) {
            if (sheetName === SHEET_NAMES.OVERVIEW) continue; // Overview is handled separately

            var sheet = spreadsheet.getSheetByName(sheetName);
            if (!sheet) {
                Logger.log("Warning: Sheet '" + sheetName + "' not found. Skipping results for this category.");
                continue;
            }

            var results = ALL_RESULTS[sheetName];
            if (results.length === 0) {
                Logger.log("No results to write for sheet: " + sheetName);
                // Clear old data if sheet exists but has no new results
                if (sheet.getLastRow() > 1) {
                   sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent().clearFormat();
                }
                continue; // Skip empty sheets
            }

            // Clear existing data below header before writing new data (redundant with initialize, but safe)
             if (sheet.getLastRow() > 1) {
                sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent().clearFormat();
            }

            var startRow = 2; // Always start writing data from row 2
            sheet.getRange(startRow, 1, results.length, results[0].length).setValues(results);
            Logger.log("Successfully wrote " + results.length + " findings to sheet: " + sheetName);

            // Conditional formatting removed due to API errors
            // Logger.log("Conditional formatting skipped for sheet '" + sheetName + "' due to API limitations.");

            // Auto-resize columns
            for (var i = 1; i <= results[0].length; i++) {
                try { // Wrap auto-resize in try-catch as well, just in case
                    sheet.autoResizeColumn(i);
                } catch (resizeError) {
                     Logger.log("Error auto-resizing column " + i + " on sheet '" + sheetName + "': " + resizeError);
                }
            }
        } // End loop through sheets

    } catch (e) {
        Logger.log("Error writing results to spreadsheet: " + e);
        if (e.message) {
            Logger.log("Error message: " + e.message);
        }
        if (e.stack) {
            Logger.log("Error stack: " + e.stack);
        }
        // Log results here if writing failed
        Logger.log("Logging results here due to spreadsheet write error:");
        for (var sheetName in ALL_RESULTS) {
             Logger.log("--- Results for " + sheetName + " ---");
             ALL_RESULTS[sheetName].forEach(function(row) {
                 Logger.log(row.join(" | "));
             });
        }
    }
}


// --- Audit Modules ---

/**
 * Audits general account settings.
 * Checklist Items: Location targeting, Language settings, Device targeting, Ad scheduling,
 *                  Auto-tagging, IP exclusions, Currency and time zone.
 */
function auditAccountSettings() {
  var category = "Account Settings";
  Logger.log("--- Auditing " + category + " ---");
  try {
    var account = AdsApp.currentAccount();

    // Currency and Time Zone (Info)
    addResult(category, "Currency and Time Zone", "Info", "Currency: " + account.getCurrencyCode() + ", Time Zone: " + account.getTimeZone(), "Verify these are correct for the business.");

    // Auto-tagging - Method unavailable, requires manual check
    addResult(category, "Auto-tagging Enabled", "Info", "Manual Check Required", "Verify auto-tagging status in Account Settings > Tracking. Enable it for proper GA tracking if disabled.");

    // IP Exclusions (Requires iterating through campaigns)
    var ipExclusionCount = 0;
    var campaignsWithIpExclusions = 0;
    var campaignIterator = AdsApp.campaigns().withCondition("Status = ENABLED").get();
    while (campaignIterator.hasNext()) {
        var campaign = campaignIterator.next();
        try { // Add try-catch around targeting access
            // Check if the excludedIpAddresses method exists before calling it
            if (campaign.targeting() && typeof campaign.targeting().excludedIpAddresses === 'function') {
                var excludedIps = campaign.targeting().excludedIpAddresses().get();
                if (excludedIps.hasNext()) {
                    campaignsWithIpExclusions++;
                    while(excludedIps.hasNext()){
                        excludedIps.next(); // Just need to count
                        ipExclusionCount++;
                    }
                }
            } else {
                 addResult(category, "IP Exclusions Check", "Info", "IP exclusion check skipped for campaign '" + campaign.getName() + "'.", "Method unavailable (likely incompatible campaign type like Performance Max).");
            }
        } catch (targetingError) {
             addResult(category, "IP Exclusions Check", "Error", "Could not check IP exclusions for campaign '" + campaign.getName() + "': " + targetingError, "Check script permissions or API changes.");
        }
    }
     addResult(category, "IP Exclusions", ipExclusionCount > 0 ? "Info" : "Warn", ipExclusionCount + " IPs excluded across " + campaignsWithIpExclusions + " checked campaigns.", ipExclusionCount > 0 ? "Review excluded IPs periodically for relevance." : "Consider adding IP exclusions for irrelevant traffic (e.g., office IPs, known bots).");


    // Location, Language, Device, Ad Scheduling require campaign-level checks
    // These are partially covered in auditCampaignOptimization but noted here for completeness.
    addResult(category, "Location Targeting", "Info", "Checked per campaign.", "Ensure targeting matches business goals in each campaign.");
    addResult(category, "Language Settings", "Info", "Checked per campaign.", "Ensure language aligns with ad content and target audience in each campaign.");
    addResult(category, "Device Targeting", "Info", "Checked per campaign.", "Review device performance and apply bid adjustments as needed per campaign.");
    addResult(category, "Ad Scheduling", "Info", "Checked per campaign.", "Review performance by time/day and apply schedules or bid adjustments per campaign.");

  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}

/**
 * Audits account structure elements like campaign organization and naming.
 * Checklist Items: Campaign organization, Ad group themes, Overlapping keywords (basic check),
 *                  Campaign naming conventions, Budget allocation.
 */
function auditAccountStructure() {
  var category = "Account Structure";
  Logger.log("--- Auditing " + category + " ---");
  var campaignsChecked = 0;
  var campaignsWithPoorNaming = 0;
  var totalBudget = 0;

  try {
    var campaignIterator = AdsApp.campaigns()
      .withCondition("Status IN [ENABLED, PAUSED]")
      .get();

    while (campaignIterator.hasNext()) {
      var campaign = campaignIterator.next();
      campaignsChecked++;
      var campaignName = campaign.getName();
      var budgetAmount = 0;
      try { // Handle potential errors getting budget (e.g., for certain campaign types)
          var budget = campaign.getBudget();
          if (budget) {
              budgetAmount = budget.getAmount();
          }
      } catch (budgetError) {
          addResult(category, "Budget Check", "Warn", "Could not retrieve budget for campaign '" + campaignName + "': " + budgetError, "Review budget setting manually.");
      }
      totalBudget += budgetAmount;


      // Campaign Naming Convention
      if (CONFIG.CAMPAIGN_NAMING_CONVENTION_REGEX && !CONFIG.CAMPAIGN_NAMING_CONVENTION_REGEX.test(campaignName)) {
        campaignsWithPoorNaming++;
        addResult(category, "Campaign Naming Convention", "Warn", "Campaign '" + campaignName + "' doesn't match pattern: " + CONFIG.CAMPAIGN_NAMING_CONVENTION_REGEX.source, "Standardize campaign naming for better organization.");
      }

      // Budget Allocation (Info - requires context)
      addResult(category, "Budget Allocation", "Info", "Campaign '" + campaignName + "' Budget: " + budgetAmount.toFixed(2) + " " + AdsApp.currentAccount().getCurrencyCode(), "Ensure budget aligns with campaign priority and performance.");

      // Ad Group Theming & Overlapping Keywords (Checked in respective modules)
    }

    if (campaignsChecked > 0 && campaignsWithPoorNaming === 0) {
       addResult(category, "Campaign Naming Convention", "Pass", "All " + campaignsChecked + " checked campaigns follow the naming convention.", "Maintain consistent naming.");
    } else if (campaignsChecked > 0 && campaignsWithPoorNaming > 0) {
        // Already logged warnings individually
    } else if (campaignsChecked === 0) {
         addResult(category, "Campaign Naming Convention", "Info", "No active/paused campaigns found to check.", "N/A");
    }

    addResult(category, "Total Account Budget", "Info", "Sum of daily budgets for checked campaigns: " + totalBudget.toFixed(2) + " " + AdsApp.currentAccount().getCurrencyCode(), "Verify total budget aligns with overall advertising goals.");
    addResult(category, "Campaign Organization", "Info", "Manual Review Required", "Review campaign goals (e.g., Search, Display, Video) and structure (e.g., by product, service, location). Ensure logical grouping.");
    addResult(category, "Ad Group Theming", "Info", "Checked in Ad Groups module.", "Ensure ad groups contain tightly themed keywords.");
    addResult(category, "Overlapping Keywords", "Info", "Basic checks in Keywords module.", "Perform deeper analysis using Search Terms Report or dedicated tools if needed.");


  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits conversion tracking setup.
 * Checklist Items: Tracking implemented, Primary/Secondary actions, Tag firing (limited check),
 *                  Duplicate tracking (limited check), Conversion values, GA linking.
 */
function auditConversionTracking() {
  var category = "Conversion Tracking";
  Logger.log("--- Auditing " + category + " ---");
  addResult(category, "Conversion Action Check", "Info", "Manual Check Required", "Checking specific conversion actions via standard Ads Scripts has limitations. Please review conversion actions, primary/secondary settings, values, and GA linking manually in the UI (Tools & Settings > Measurement > Conversions).");
  /* // Commenting out due to AdsApp.conversionActions() limitation.
  try {
    // var conversionActions = AdsApp.conversionActions().get(); // This method is unavailable
    // --- Existing code commented out ---
    var primaryActions = 0;
    var secondaryActions = 0;
    var actionsWithValue = 0;
    var totalActions = 0;
    var tagManagerActions = 0;
    var analyticsGoals = 0;
    var websiteActions = 0;
    var duplicateCheckMap = {}; // Simple check for identical names

    if (!conversionActions.hasNext()) {
      addResult(category, "Conversion Tracking Implemented", "Fail", "No conversion actions found in the account.", "Implement conversion tracking immediately.");
      return; // Stop if no actions exist
    }

     addResult(category, "Conversion Tracking Implemented", "Pass", "Conversion actions found.", "Review specific actions below.");

    while (conversionActions.hasNext()) {
      var action = conversionActions.next();
      totalActions++;
      var actionName = action.getName();
      var actionCategory = action.getCategory();
      var actionStatus = action.getStatus(); // ENABLED, REMOVED, HIDDEN
      var actionOrigin = 'N/A'; // Default
      var actionCountingType = 'N/A'; // Default
      var actionValueSettings = null; // Default

      // Use try-catch for methods that might not exist on all action types/versions
      try { actionOrigin = action.getOrigin ? action.getOrigin() : 'N/A'; } catch (e) { Logger.log("Warning: Could not get origin for action '" + actionName + "'."); }
      try { actionCountingType = action.getCountingType ? action.getCountingType() : 'N/A'; } catch (e) { Logger.log("Warning: Could not get counting type for action '" + actionName + "'."); }
      try { actionValueSettings = action.getValueSettings ? action.getValueSettings() : null; } catch (e) { Logger.log("Warning: Could not get value settings for action '" + actionName + "'."); }


      if (actionStatus !== 'ENABLED') {
          addResult(category, "Action Status", "Warn", "Action '" + actionName + "' is " + actionStatus, "Review if this action should be enabled or removed.");
          continue; // Skip checks for non-enabled actions
      }

      // Primary/Secondary Approximation
      if (['PURCHASE', 'LEAD', 'SIGN_UP', 'PAGE_VIEW', 'OUTBOUND_CLICK', 'PHONE_CALL_LEAD', 'SUBMIT_LEAD_FORM', 'BOOK_APPOINTMENT', 'REQUEST_QUOTE'].indexOf(actionCategory) !== -1) {
          primaryActions++;
          addResult(category, "Primary Conversion Actions", "Info", "Action '" + actionName + "' (Category: " + actionCategory + ") likely primary.", "Verify this action is correctly set as primary in the UI if applicable.");
      } else {
          secondaryActions++;
           addResult(category, "Secondary Conversion Actions", "Info", "Action '" + actionName + "' (Category: " + actionCategory + ") likely secondary.", "Verify this action is correctly set as secondary in the UI if applicable.");
      }

      // Conversion Values
      var hasValue = false;
      if (actionValueSettings) {
          try {
              if (actionValueSettings.getDefaultValue() > 0 || actionValueSettings.getAlwaysUseDefaultValue() === false) {
                  actionsWithValue++;
                  hasValue = true;
              }
          } catch (vsError) {
               Logger.log("Warning: Could not fully check value settings for action '" + actionName + "'. " + vsError);
          }
      }
      addResult(category, "Conversion Values Assigned", hasValue ? "Pass" : "Warn", "Action '" + actionName + "' " + (hasValue ? "has value settings." : "does not seem to have specific value settings."), hasValue ? "Ensure values are accurate." : "Assign conversion values if applicable (e.g., for purchases, leads with estimated value).");


      // Tag Firing (API Limitation)
      if (actionOrigin === 'WEBSITE' || actionOrigin === 'GOOGLE_TAG_MANAGER') websiteActions++;
      if (actionOrigin === 'GOOGLE_ANALYTICS') analyticsGoals++;
      if (actionOrigin === 'GOOGLE_TAG_MANAGER') tagManagerActions++;

      addResult(category, "Tag Firing Check", "Info", "Action '" + actionName + "' Origin: " + actionOrigin + ", Status: " + actionStatus, "API cannot confirm live tag firing. Use Google Tag Assistant or check recent conversion data manually.");

      // Duplicate Tracking (Basic Name Check)
      if (duplicateCheckMap[actionName]) {
        addResult(category, "Potential Duplicate Tracking", "Warn", "Multiple enabled conversion actions found with the name '" + actionName + "'.", "Investigate if these are duplicates or intentionally named similarly. Ensure correct counting settings.");
      }
      duplicateCheckMap[actionName] = true;

    }

    if (primaryActions === 0 && totalActions > 0) {
        addResult(category, "Primary Conversion Actions Defined", "Fail", "No clear primary conversion actions identified among enabled actions.", "Define at least one primary conversion action representing key business goals.");
    } else if (primaryActions > 0) {
         addResult(category, "Primary Conversion Actions Defined", "Pass", primaryActions + " potential primary action(s) identified.", "Verify in UI.");
    }

    if (secondaryActions > 0) {
        addResult(category, "Secondary Conversion Actions Tracked", "Pass", secondaryActions + " potential secondary action(s) identified.", "Verify in UI.");
    } else {
         addResult(category, "Secondary Conversion Actions Tracked", "Info", "No clear secondary actions identified.", "Consider tracking micro-conversions as secondary actions if valuable.");
    }


    // Google Analytics Linking (API Limitation)
    addResult(category, "Google Analytics Linked", "Info", "Manual Check Required", "Verify in Google Ads UI (Tools & Settings > Linked Accounts > Google Analytics) that the correct GA property is linked.");
    if (analyticsGoals > 0) {
        addResult(category, "GA Goals Imported", "Pass", analyticsGoals + " conversion action(s) sourced from Google Analytics found.", "Ensure imported goals are relevant and correctly configured in GA.");
    } else {
        addResult(category, "GA Goals Imported", "Warn", "No conversion actions sourced from Google Analytics found.", "If using GA goals, ensure they are imported into Google Ads.");
    }

  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
  */
}


/**
 * Audits keywords: match types, negatives, search terms (review note), duplicates, performance.
 * Checklist Items: Alignment with goals, Match types, Negative keywords, Search terms review,
 *                  Duplicate keywords, Low performers, Keyword bids.
 */
function auditKeywords() {
  var category = "Keywords";
  Logger.log("--- Auditing " + category + " ---");
  var keywordsChecked = 0;
  var lowQsKeywords = 0;
  var broadMatchKeywords = 0;
  var phraseMatchKeywords = 0;
  var exactMatchKeywords = 0;
  var negativeKeywordsAccount = 0;
  var negativeKeywordsCampaign = 0;
  var negativeKeywordsAdGroup = 0;
  // Removed lowPerformingKeywords as stats are unavailable
  var duplicateKeywordMap = {}; // Store keyword text + match type + ad group ID + campaign ID
  var statsErrorCount = 0; // Count keywords where stats failed
  var qsErrorCount = 0; // Count keywords where QS failed

  try {
    // Account Level Negatives
    var accountNegativeLists = AdsApp.negativeKeywordLists().get();
    while(accountNegativeLists.hasNext()){
        var list = accountNegativeLists.next();
        negativeKeywordsAccount += list.negativeKeywords().get().totalNumEntities();
    }
     addResult(category, "Account Negative Keywords", negativeKeywordsAccount > 0 ? "Pass" : "Warn", negativeKeywordsAccount + " negatives found in lists.", negativeKeywordsAccount > 0 ? "Review lists periodically." : "Consider creating account-level negative lists for universally irrelevant terms.");


    var campaignIterator = AdsApp.campaigns().withCondition("Status = ENABLED").get();
    while (campaignIterator.hasNext()) {
      var campaign = campaignIterator.next();
      var campaignName = campaign.getName();
      var campaignId = campaign.getId();

      // Campaign Level Negatives
      var campaignNegatives = 0;
      try {
          campaignNegatives = campaign.negativeKeywords().get().totalNumEntities();
          negativeKeywordsCampaign += campaignNegatives;
          if (campaignNegatives === 0) {
             addResult(category, "Campaign Negative Keywords", "Warn", "Campaign '" + campaignName + "' has no direct negative keywords.", "Add campaign-level negatives relevant to this campaign (or ensure coverage via lists).");
          }
      } catch (negError) {
           addResult(category, "Campaign Negative Keywords Check", "Error", "Could not check negatives for campaign '" + campaignName + "': " + negError, "Check campaign type/permissions.");
      }


      var adGroupIterator = campaign.adGroups().withCondition("Status = ENABLED").get();
      while (adGroupIterator.hasNext()) {
        var adGroup = adGroupIterator.next();
        var adGroupName = adGroup.getName();
        var adGroupId = adGroup.getId();

        // Ad Group Level Negatives
        var adGroupNegatives = 0;
        try {
            adGroupNegatives = adGroup.negativeKeywords().get().totalNumEntities();
            negativeKeywordsAdGroup += adGroupNegatives;
             if (adGroupNegatives === 0) {
                addResult(category, "Ad Group Negative Keywords", "Warn", "Ad Group '" + adGroupName + "' in Campaign '" + campaignName + "' has no negative keywords.", "Add ad group-level negatives for fine-tuning.");
            }
        } catch (negAgError) {
             addResult(category, "Ad Group Negative Keywords Check", "Error", "Could not check negatives for ad group '" + adGroupName + "': " + negAgError, "Check permissions.");
        }


        var keywordIterator = adGroup.keywords()
          .withCondition("Status = ENABLED")
          //.withCondition("KeywordMatchType != BROAD") // Let's check all enabled for performance
          // .forDateRange("LAST_30_DAYS") // Date range not needed if not getting stats
          .get();

        while (keywordIterator.hasNext()) {
          var keyword = keywordIterator.next();
          keywordsChecked++;
          var keywordText = keyword.getText();
          var matchType = keyword.getMatchType();
          var qs = null;

          // Skip stats retrieval entirely as keyword.getStats() is unavailable
          statsErrorCount++;

          // Try getting Quality Score
          try {
              qs = keyword.getQualityScore();
          } catch (qsError) {
               qsErrorCount++;
               Logger.log("Warning: Could not get QS for keyword '" + keywordText + "' in Ad Group '" + adGroupName + "'. Error: " + qsError);
               // Don't log to sheet here, handled in summary
          }


          // Match Type Distribution
          if (matchType === "BROAD") broadMatchKeywords++;
          else if (matchType === "PHRASE") phraseMatchKeywords++;
          else if (matchType === "EXACT") exactMatchKeywords++;

          // Quality Score Check (if available)
          if (qs !== null && qs < CONFIG.MIN_QUALITY_SCORE) {
            lowQsKeywords++;
            addResult(category, "Low Quality Score", "Fail", "Keyword '" + keywordText + "' (" + matchType + ") in Ad Group '" + adGroupName + "' has QS: " + qs, "Improve ad relevance, expected CTR, or landing page experience.");
          }

          // Low Performance Check - Removed due to getStats() unavailability

          // Duplicate Keyword Check (Across Ad Groups within the same Campaign - basic version)
          var duplicateKey = campaignId + ":" + keywordText + ":" + matchType;
          if (duplicateKeywordMap[duplicateKey] && duplicateKeywordMap[duplicateKey] !== adGroupId) {
             addResult(category, "Potential Duplicate Keyword (Cross-AdGroup)", "Warn", "Keyword '" + keywordText + "' (" + matchType + ") found in Ad Group '" + adGroupName + "' and also potentially in Ad Group ID '" + duplicateKeywordMap[duplicateKey] + "' within Campaign '" + campaignName + "'.", "Ensure keywords don't compete across ad groups within the same campaign unless intended (e.g., different geo-targets or match types).");
          }
          // Store the first ad group ID found for this keyword in this campaign
          if (!duplicateKeywordMap[duplicateKey]) {
              duplicateKeywordMap[duplicateKey] = adGroupId;
          }


          // Keyword Bids (Info - depends on bidding strategy)
          var bid = "Managed by Strategy";
          try {
              bid = keyword.bidding().getCpc() || adGroup.bidding().getCpc() || "Managed by Strategy";
          } catch (bidError) { /* Ignore if bidding info not available */ }
          addResult(category, "Keyword Bids", "Info", "Keyword '" + keywordText + "' (" + matchType + ") Bid: " + (typeof bid === 'number' ? bid.toFixed(2) : bid), "Ensure bids align with performance and bidding strategy.");

        } // End keyword loop
      } // End ad group loop
    } // End campaign loop

    // Summary Results
    addResult(category, "Keyword Match Types", "Info", "Broad: " + broadMatchKeywords + ", Phrase: " + phraseMatchKeywords + ", Exact: " + exactMatchKeywords + " (Total Enabled Checked: " + keywordsChecked + ")", "Ensure match type usage aligns with campaign goals (e.g., control vs. reach). Review broad match performance carefully.");
    addResult(category, "Total Negative Keywords", "Info", "Account Lists: " + negativeKeywordsAccount + ", Campaign Level: " + negativeKeywordsCampaign + ", Ad Group Level: " + negativeKeywordsAdGroup, "Ensure comprehensive negative keyword coverage at appropriate levels.");
    addResult(category, "Search Terms Report Review", "Info", "Manual Check Required", "Regularly review the Search Terms Report in the UI to find new positive and negative keyword opportunities.");
    addResult(category, "Keyword Alignment with Goals", "Info", "Manual Review Required", "Ensure keywords in each ad group are relevant to the ad copy, landing page, and overall campaign objective.");

    if (keywordsChecked > 0 && lowQsKeywords === 0) {
        addResult(category, "Low Quality Score", "Pass", "No keywords found below QS " + CONFIG.MIN_QUALITY_SCORE + " (out of " + keywordsChecked + " checked with QS data).", "Maintain good QS practices.");
    }
    // Add summary note about skipped performance checks
    if (statsErrorCount > 0) {
         addResult(category, "Keyword Performance Checks", "Info", "Performance checks (CTR, CPA, etc.) skipped for " + statsErrorCount + " keywords due to API limitations (getStats unavailable).", "Review keyword performance manually using Google Ads reports.");
    }
     if (qsErrorCount > 0) {
         addResult(category, "Quality Score Errors", "Warn", "Could not retrieve QS for " + qsErrorCount + " keywords.", "Check script logs for details. May indicate permission issues or API changes.");
    }


  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits ad groups: keyword count, ad count, bids, naming.
 * Checklist Items: Keyword count, Ad count, Ad group bids, Competing keywords (note), Naming.
 */
function auditAdGroups() {
  var category = "Ad Groups";
  Logger.log("--- Auditing " + category + " ---");
  var adGroupsChecked = 0;
  var adGroupsWithLowAdCount = 0;
  var adGroupsWithHighKeywordCount = 0;
  var adGroupsWithPoorNaming = 0;

  try {
    var adGroupIterator = AdsApp.adGroups()
      .withCondition("Status = ENABLED")
      .withCondition("CampaignStatus = ENABLED")
      .get();

    while (adGroupIterator.hasNext()) {
      var adGroup = adGroupIterator.next();
      adGroupsChecked++;
      var adGroupName = adGroup.getName();
      var campaignName = adGroup.getCampaign().getName();

      // Keyword Count
      var keywordCount = 0;
      try {
          keywordCount = adGroup.keywords().withCondition("Status = ENABLED").get().totalNumEntities();
          if (keywordCount > CONFIG.MAX_KEYWORDS_PER_ADGROUP) {
            adGroupsWithHighKeywordCount++;
            addResult(category, "Keyword Count", "Warn", "Ad Group '" + adGroupName + "' (" + campaignName + ") has " + keywordCount + " keywords (>" + CONFIG.MAX_KEYWORDS_PER_ADGROUP + ").", "Consider splitting into more tightly themed ad groups for better relevance.");
          } else if (keywordCount === 0) {
             // Only warn if it's not a DSA ad group
             if (adGroup.getCampaign().getAdvertisingChannelType() !== 'SEARCH' || !adGroup.isDynamic()) {
                 addResult(category, "Keyword Count", "Warn", "Ad Group '" + adGroupName + "' (" + campaignName + ") has 0 enabled keywords.", "Add relevant keywords or pause the ad group if it's not needed (and not DSA).");
             } else {
                  addResult(category, "Keyword Count", "Info", "Ad Group '" + adGroupName + "' (" + campaignName + ") is DSA and has 0 keywords.", "Expected for DSA ad groups.");
             }
          }
      } catch (kwCountError) {
           addResult(category, "Keyword Count Check", "Error", "Could not check keyword count for ad group '" + adGroupName + "': " + kwCountError, "Check permissions.");
      }


      // Ad Count
      var adCount = 0;
      try {
          adCount = adGroup.ads().withCondition("Status = ENABLED").get().totalNumEntities();
          if (adCount < CONFIG.MIN_ADS_PER_ADGROUP) {
            adGroupsWithLowAdCount++;
            addResult(category, "Ad Count", "Fail", "Ad Group '" + adGroupName + "' (" + campaignName + ") has " + adCount + " enabled ads (<" + CONFIG.MIN_ADS_PER_ADGROUP + ").", "Create at least " + CONFIG.MIN_ADS_PER_ADGROUP + " relevant ads per ad group for testing and optimization.");
          }
      } catch (adCountError) {
           addResult(category, "Ad Count Check", "Error", "Could not check ad count for ad group '" + adGroupName + "': " + adCountError, "Check permissions.");
      }


      // Ad Group Naming
      if (CONFIG.ADGROUP_NAMING_CONVENTION_REGEX && !CONFIG.ADGROUP_NAMING_CONVENTION_REGEX.test(adGroupName)) {
        adGroupsWithPoorNaming++;
        addResult(category, "Ad Group Naming", "Warn", "Ad Group '" + adGroupName + "' (" + campaignName + ") doesn't match pattern: " + CONFIG.ADGROUP_NAMING_CONVENTION_REGEX.source, "Standardize ad group naming.");
      }

      // Ad Group Bids (Info - depends on strategy)
      var bid = "Managed by Strategy";
      try {
          bid = adGroup.bidding().getCpc() || "Managed by Strategy";
      } catch (bidError) { /* Ignore */ }
      addResult(category, "Ad Group Bids", "Info", "Ad Group '" + adGroupName + "' (" + campaignName + ") Bid: " + (typeof bid === 'number' ? bid.toFixed(2) : bid), "Ensure bids align with performance goals and bidding strategy.");

      // Competing Keywords within Ad Group (Checked in Keywords module - duplicates)
    }

    // Summary Results
    if (adGroupsChecked > 0) {
        if (adGroupsWithLowAdCount === 0) addResult(category, "Ad Count", "Pass", "All " + adGroupsChecked + " checked ad groups have at least " + CONFIG.MIN_ADS_PER_ADGROUP + " ads.", "Continue A/B testing ads.");
        if (adGroupsWithHighKeywordCount === 0) addResult(category, "Keyword Count", "Pass", "All " + adGroupsChecked + " checked ad groups have a reasonable number of keywords (<= " + CONFIG.MAX_KEYWORDS_PER_ADGROUP + ").", "Maintain tight keyword themes.");
        if (adGroupsWithPoorNaming === 0) addResult(category, "Ad Group Naming", "Pass", "All " + adGroupsChecked + " checked ad groups follow the naming convention.", "Maintain consistent naming.");
    } else {
        addResult(category, "General Check", "Info", "No enabled ad groups found in enabled campaigns.", "N/A");
    }
     addResult(category, "Competing Keywords within Ad Group", "Info", "Basic duplicate checks in Keywords module.", "Manually review keyword themes within ad groups to ensure they aren't competing semantically.");


  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits ad copy: RSA usage, policy status. Subjective checks are manual.
 * Checklist Items: Headlines/Descriptions (manual), CTAs (manual), Tailored to theme (manual),
 *                  RSA usage, Ad variations (manual), Spelling/Grammar (manual), Policy compliance.
 */
function auditAdCopy() {
  var category = "Ad Copy";
  Logger.log("--- Auditing " + category + " ---");
  var adsChecked = 0;
  var rsaCount = 0;
  var nonRsaSearchAds = 0; // Expanded Text Ads (ETAs)
  var disapprovedAds = 0;
  var adGroupsWithNoRsa = {}; // Track ad groups lacking RSAs {adGroupId: {name: agName, campaign: campName}}

  try {
    var adIterator = AdsApp.ads()
      .withCondition("Status = ENABLED")
      .withCondition("AdGroupStatus = ENABLED")
      .withCondition("CampaignStatus = ENABLED")
      // .withCondition("Type IN [RESPONSIVE_SEARCH_AD, EXPANDED_TEXT_AD]") // Check all relevant types
      .get();

     while (adIterator.hasNext()) {
        var ad = adIterator.next();
        adsChecked++;
        var adType = ad.getType();
        var policyApprovalStatus = "UNKNOWN";
        var adGroupId = ad.getAdGroup().getId();
        var adGroupName = ad.getAdGroup().getName();
        var campaignName = ad.getCampaign().getName();
        var isSearchCampaign = ad.getCampaign().getAdvertisingChannelType() === "SEARCH";

        try { policyApprovalStatus = ad.getPolicyApprovalStatus(); } catch (e) { Logger.log("Warning: Could not get policy status for ad in " + adGroupName); }


        if (adType === "RESPONSIVE_SEARCH_AD") {
            rsaCount++;
            adGroupsWithNoRsa[adGroupId] = false; // Mark as having an RSA
        } else if (adType === "EXPANDED_TEXT_AD") {
            nonRsaSearchAds++;
             if (isSearchCampaign && adGroupsWithNoRsa[adGroupId] !== false) { // If not already marked as having RSA in a Search campaign
                 adGroupsWithNoRsa[adGroupId] = { name: adGroupName, campaign: campaignName };
            }
        } else {
             // Other ad types might exist (e.g., App, Call Only) - consider if RSA check applies only to Search
             if (isSearchCampaign && adGroupsWithNoRsa[adGroupId] !== false) {
                 adGroupsWithNoRsa[adGroupId] = { name: adGroupName, campaign: campaignName };
             }
        }

        // Policy Compliance
        if (policyApprovalStatus === "DISAPPROVED") {
            disapprovedAds++;
            var policyTopics = "N/A";
            try {
                 policyTopics = ad.getPolicyTopics ? ad.getPolicyTopics().map(function(topic){ return topic.getPolicyTopicType(); }).join(', ') : "N/A";
            } catch(policyError) { /* Ignore if method doesn't exist */ }

            addResult(category, "Policy Compliance", "Fail", "Ad in Ad Group '" + adGroupName + "' (" + campaignName + ") is DISAPPROVED. Topics: " + policyTopics, "Review policy violations and edit or remove the ad.");
        } else if (policyApprovalStatus !== "APPROVED" && policyApprovalStatus !== "UNKNOWN") {
             addResult(category, "Policy Compliance", "Warn", "Ad in Ad Group '" + adGroupName + "' (" + campaignName + ") status is " + policyApprovalStatus, "Monitor status; may require action if it becomes disapproved.");
        }
     }

     // Summarize RSA usage
     if (adsChecked > 0) {
         if (rsaCount > 0) {
             addResult(category, "Responsive Search Ads (RSAs) Utilized", "Pass", rsaCount + " enabled RSAs found.", "Ensure RSAs have sufficient high-quality assets.");
         }
         if (nonRsaSearchAds > 0) {
             addResult(category, "Legacy Ad Formats (ETAs)", "Warn", nonRsaSearchAds + " enabled Expanded Text Ads found.", "Consider migrating ETAs to RSAs as ETAs can no longer be created or edited.");
         }

         var missingRsaCount = 0;
         for (var agId in adGroupsWithNoRsa) {
             if (adGroupsWithNoRsa[agId]) { // If it's still an object (meaning no RSA was found for this Search AG)
                 missingRsaCount++;
                 addResult(category, "RSA Usage per Ad Group", "Fail", "Search Ad Group '" + adGroupsWithNoRsa[agId].name + "' (" + adGroupsWithNoRsa[agId].campaign + ") appears to be missing an enabled RSA.", "Ensure each active Search ad group has at least one enabled RSA.");
             }
         }
         if (missingRsaCount === 0 && rsaCount > 0) { // Check if any RSAs were found at all
              addResult(category, "RSA Usage per Ad Group", "Pass", "All checked Search ad groups with ads appear to have at least one RSA.", "Good.");
         } else if (missingRsaCount === 0 && rsaCount === 0 && nonRsaSearchAds === 0) {
             // No RSAs, but also no other Search ads found to necessitate them
             addResult(category, "RSA Usage per Ad Group", "Info", "No enabled Search ads (RSA or ETA) found to assess RSA coverage.", "N/A");
         }

     } else {
         addResult(category, "General Check", "Info", "No enabled ads found in enabled ad groups/campaigns.", "N/A");
     }

     if (adsChecked > 0 && disapprovedAds === 0) {
         addResult(category, "Policy Compliance", "Pass", "No disapproved ads found among " + adsChecked + " checked ads.", "Maintain policy compliance.");
     }

     // Manual Checks
     addResult(category, "Headlines & Descriptions", "Info", "Manual Review Required", "Review RSAs for compelling headlines/descriptions, asset variety, and strength indicators (Pinning usage, Asset performance).");
     addResult(category, "Call-to-Actions (CTAs)", "Info", "Manual Review Required", "Ensure ads include clear and strong CTAs relevant to the offering.");
     addResult(category, "Ad Tailoring to Theme", "Info", "Manual Review Required", "Verify that ad copy (especially RSAs) is highly relevant to the keywords within the ad group.");
     addResult(category, "Ad Variations (A/B Testing)", "Info", "Manual Review Required", "Regularly test different ad copy elements (headlines, descriptions, CTAs) using experiments or by monitoring RSA asset performance.");
     addResult(category, "Spelling & Grammar", "Info", "Manual Review Required", "Proofread all ad copy for errors.");


  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits ad extensions (assets): checks for presence and basic status.
 * Checklist Items: Sitelinks, Callouts, Structured Snippets, Call, Location, Price.
 */
function auditAdExtensions() {
  // Reverting to older .extensions() methods as .assets() seem unavailable
  var category = "Ad Extensions (Assets)";
  Logger.log("--- Auditing " + category + " ---");
  var extensionsChecked = {
    SITELINK: { count: 0, campaigns: 0, adgroups: 0 },
    CALLOUT: { count: 0, campaigns: 0, adgroups: 0 },
    STRUCTURED_SNIPPET: { count: 0, campaigns: 0, adgroups: 0 },
    CALL: { count: 0, campaigns: 0, adgroups: 0 }, // Using phoneNumbers()
    LOCATION: { count: 0, campaigns: 0, adgroups: 0 }, // Manual check needed
    PRICE: { count: 0, campaigns: 0, adgroups: 0 },
  };
  var campaignsWithoutSitelinks = 0;
  var campaignsWithoutCallouts = 0;
  var campaignsChecked = 0;

  try {
    // Account Level Extensions (Check if methods exist)
    var accountExtensions = AdsApp.currentAccount().extensions();
    try {
        var accCallouts = accountExtensions.callouts().get();
        while(accCallouts.hasNext()){ extensionsChecked.CALLOUT.count++; accCallouts.next(); }
        addResult(category, "Account Level Callouts", "Info", extensionsChecked.CALLOUT.count + " callout extensions found at account level.", "Ensure these are appropriate account-wide.");
    } catch(e) { addResult(category, "Account Level Callouts Check", "Warn", "Could not check account-level callouts: " + e, "Method might be unavailable."); }
    // Add similar checks for Sitelinks, Snippets etc. if needed

    // Location Extension Check (Manual)
    addResult(category, "Location Extensions", "Info", "Manual Check Required", "Verify Google Business Profile is linked at the account or campaign level in the UI (Extensions/Assets section).");


    // Campaign & Ad Group Level Extensions
    var campaignIterator = AdsApp.campaigns()
      .withCondition("Status = ENABLED")
      .get();

    while (campaignIterator.hasNext()) {
      var campaign = campaignIterator.next();
      campaignsChecked++;
      var campaignName = campaign.getName();
      var hasSitelinks = false;
      var hasCallouts = false;

      // Check Campaign Level Extensions
      try {
          var campaignExtensions = campaign.extensions();
          // Check if methods exist before calling
          if (campaignExtensions && typeof campaignExtensions.sitelinks === 'function') {
              var campSitelinks = campaignExtensions.sitelinks().get(); hasSitelinks = campSitelinks.hasNext(); extensionsChecked.SITELINK.campaigns += hasSitelinks ? 1:0; while(campSitelinks.hasNext()){ extensionsChecked.SITELINK.count++; campSitelinks.next();}
          }
          if (campaignExtensions && typeof campaignExtensions.callouts === 'function') {
              var campCallouts = campaignExtensions.callouts().get(); hasCallouts = campCallouts.hasNext(); extensionsChecked.CALLOUT.campaigns += hasCallouts ? 1:0; while(campCallouts.hasNext()){ extensionsChecked.CALLOUT.count++; campCallouts.next();}
          }
          if (campaignExtensions && typeof campaignExtensions.structuredSnippets === 'function') {
              var campSnippets = campaignExtensions.structuredSnippets().get(); extensionsChecked.STRUCTURED_SNIPPET.campaigns += campSnippets.hasNext() ? 1:0; while(campSnippets.hasNext()){ extensionsChecked.STRUCTURED_SNIPPET.count++; campSnippets.next();}
          } else if (campaignExtensions) { // Log info only if extensions object exists but method doesn't
               addResult(category, "Campaign Structured Snippet Check", "Info", "Structured Snippet check skipped for campaign '" + campaignName + "'.", "Method unavailable (likely incompatible campaign type or API change).");
          }
          if (campaignExtensions && typeof campaignExtensions.phoneNumbers === 'function') {
              var campCalls = campaignExtensions.phoneNumbers().get(); extensionsChecked.CALL.campaigns += campCalls.hasNext() ? 1:0; while(campCalls.hasNext()){ extensionsChecked.CALL.count++; campCalls.next();}
          }
           if (campaignExtensions && typeof campaignExtensions.prices === 'function') {
              var campPrices = campaignExtensions.prices().get(); extensionsChecked.PRICE.campaigns += campPrices.hasNext() ? 1:0; while(campPrices.hasNext()){ extensionsChecked.PRICE.count++; campPrices.next();}
          }
      } catch (campExtError) {
           addResult(category, "Campaign Extension Check", "Error", "Could not check extensions for campaign '" + campaignName + "': " + campExtError, "Check script permissions or API changes.");
      }


      var adGroupIterator = campaign.adGroups().withCondition("Status = ENABLED").get();
      while (adGroupIterator.hasNext()) {
        var adGroup = adGroupIterator.next();
        // Check Ad Group Level Extensions
        try {
            var adGroupExtensions = adGroup.extensions();
             // Check if methods exist before calling
            if (adGroupExtensions && typeof adGroupExtensions.sitelinks === 'function') {
                var agSitelinks = adGroupExtensions.sitelinks().get(); if(agSitelinks.hasNext()) hasSitelinks = true; extensionsChecked.SITELINK.adgroups += agSitelinks.hasNext() ? 1:0; while(agSitelinks.hasNext()){ extensionsChecked.SITELINK.count++; agSitelinks.next();}
            }
             if (adGroupExtensions && typeof adGroupExtensions.callouts === 'function') {
                var agCallouts = adGroupExtensions.callouts().get(); if(agCallouts.hasNext()) hasCallouts = true; extensionsChecked.CALLOUT.adgroups += agCallouts.hasNext() ? 1:0; while(agCallouts.hasNext()){ extensionsChecked.CALLOUT.count++; agCallouts.next();}
            }
            if (adGroupExtensions && typeof adGroupExtensions.structuredSnippets === 'function') {
                var agSnippets = adGroupExtensions.structuredSnippets().get(); extensionsChecked.STRUCTURED_SNIPPET.adgroups += agSnippets.hasNext() ? 1:0; while(agSnippets.hasNext()){ extensionsChecked.STRUCTURED_SNIPPET.count++; agSnippets.next();}
            } else if (adGroupExtensions) {
                 // Log info only if extensions object exists but method doesn't
                 addResult(category, "AdGroup Structured Snippet Check", "Info", "Structured Snippet check skipped for ad group '" + adGroup.getName() + "'.", "Method unavailable.");
            }
             if (adGroupExtensions && typeof adGroupExtensions.phoneNumbers === 'function') {
                var agCalls = adGroupExtensions.phoneNumbers().get(); extensionsChecked.CALL.adgroups += agCalls.hasNext() ? 1:0; while(agCalls.hasNext()){ extensionsChecked.CALL.count++; agCalls.next();}
            }
             if (adGroupExtensions && typeof adGroupExtensions.prices === 'function') {
                var agPrices = adGroupExtensions.prices().get(); extensionsChecked.PRICE.adgroups += agPrices.hasNext() ? 1:0; while(agPrices.hasNext()){ extensionsChecked.PRICE.count++; agPrices.next();}
            }
        } catch (agExtError) {
             addResult(category, "AdGroup Extension Check", "Error", "Could not check extensions for ad group '" + adGroup.getName() + "': " + agExtError, "Check script permissions or API changes.");
        }
      } // End ad group loop

      // Check if campaign lacks key extensions after checking campaign and ad group levels
      if (!hasSitelinks) {
        campaignsWithoutSitelinks++;
        addResult(category, "Sitelink Extensions", "Fail", "Campaign '" + campaignName + "' appears to have no Sitelinks at campaign or ad group level.", "Add relevant Sitelinks to improve ad visibility and CTR.");
      }
       if (!hasCallouts) {
        campaignsWithoutCallouts++;
        addResult(category, "Callout Extensions", "Fail", "Campaign '" + campaignName + "' appears to have no Callouts at campaign or ad group level.", "Add Callouts highlighting key benefits or features.");
      }
    } // End campaign loop

    // Summary Results (using counts accumulated from campaign/adgroup checks)
    if (campaignsChecked > 0) {
        if (campaignsWithoutSitelinks === 0) addResult(category, "Sitelink Extensions Linked", "Pass", "All " + campaignsChecked + " checked campaigns appear to have Sitelinks.", "Ensure Sitelinks are relevant and up-to-date.");
        if (campaignsWithoutCallouts === 0) addResult(category, "Callout Extensions Linked", "Pass", "All " + campaignsChecked + " checked campaigns appear to have Callouts.", "Ensure Callouts are relevant and compelling.");
    } else {
         addResult(category, "General Check", "Info", "No enabled campaigns found to check for extensions.", "N/A");
    }

    // Summarize total counts found (Note: these count instances, not unique extensions)
    addResult(category, "Total Sitelink Instances", "Info", extensionsChecked.SITELINK.count + " instances found across " + extensionsChecked.SITELINK.campaigns + " campaigns/" + extensionsChecked.SITELINK.adgroups + " ad groups.", "Review relevance.");
    addResult(category, "Total Callout Instances", "Info", extensionsChecked.CALLOUT.count + " instances found across " + extensionsChecked.CALLOUT.campaigns + " campaigns/" + extensionsChecked.CALLOUT.adgroups + " ad groups.", "Review relevance.");
    addResult(category, "Structured Snippets", extensionsChecked.STRUCTURED_SNIPPET.count > 0 ? "Pass" : "Warn", extensionsChecked.STRUCTURED_SNIPPET.count + " instances found.", extensionsChecked.STRUCTURED_SNIPPET.count > 0 ? "Ensure relevance." : "Utilize Structured Snippets where applicable.");
    addResult(category, "Call Extensions", extensionsChecked.CALL.count > 0 ? "Pass" : "Warn", extensionsChecked.CALL.count + " instances found.", extensionsChecked.CALL.count > 0 ? "Ensure number is correct." : "Add Call Extensions if receiving calls is a goal.");
    addResult(category, "Price Extensions", extensionsChecked.PRICE.count > 0 ? "Pass" : "Warn", extensionsChecked.PRICE.count + " instances found.", extensionsChecked.PRICE.count > 0 ? "Keep prices up-to-date." : "Use Price Extensions if relevant.");
    addResult(category, "Extension Relevance & Updates", "Info", "Manual Review Required", "Periodically review all active extensions to ensure they are still relevant, accurate, and link to working URLs (for Sitelinks).");

  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits bidding strategies and budget utilization.
 * Checklist Items: Strategy alignment, Manual bidding justification, Smart bidding optimization,
 *                  Bid adjustments, Budget constraints, Impression share loss (Budget/Rank).
 */
function auditBiddingStrategies() {
  var category = "Bidding Strategies";
  Logger.log("--- Auditing " + category + " ---");
  var campaignsChecked = 0;
  var manualCpcCampaigns = 0;
  var smartBiddingCampaigns = 0; // tCPA, tROAS, Maximize Conversions, Maximize Conversion Value
  var otherBiddingCampaigns = 0; // Max Clicks, Target Impression Share, etc.
  var limitedByBudgetCampaigns = 0;
  var highIsLostRankCampaigns = 0;
  var highIsLostBudgetCampaigns = 0;
  var statsErrorCount = 0; // Count campaigns where stats failed

  try {
    var campaignIterator = AdsApp.campaigns()
      .withCondition("Status = ENABLED")
      .forDateRange("LAST_30_DAYS") // For IS stats
      .get();

    while (campaignIterator.hasNext()) {
      var campaign = campaignIterator.next();
      campaignsChecked++;
      var campaignName = campaign.getName();
      var biddingStrategyType = "UNKNOWN";
      var stats = null;
      var isLostBudget = 0;
      var isLostRank = 0;

      try { biddingStrategyType = campaign.getBiddingStrategyType(); } catch (e) { Logger.log("Warning: Could not get bidding strategy for " + campaignName); }

      // Try getting stats, but don't rely on them for core logic if they fail
      try {
          // Check if getStats method exists before calling
          if (typeof campaign.getStats === 'function') {
              stats = campaign.getStats();
              if (stats) {
                  try { isLostBudget = stats.getSearchBudgetLostImpressionShare(); } catch (e) { /* Ignore */ }
                  try { isLostRank = stats.getSearchRankLostImpressionShare(); } catch (e) { /* Ignore */ }
              }
          } else {
               statsErrorCount++;
               Logger.log("Warning: campaign.getStats() method unavailable for campaign " + campaignName);
          }
      } catch (e) {
          statsErrorCount++;
          Logger.log("Warning: Could not get stats for campaign " + campaignName + " in Bidding Strategies. Error: " + e);
          // Don't add result here, summarize at the end
      }


      // Strategy Type Count
      if (biddingStrategyType === "MANUAL_CPC" || biddingStrategyType === "MANUAL_CPM" || biddingStrategyType === "MANUAL_CPV" || biddingStrategyType === "ENHANCED_CPC") {
        manualCpcCampaigns++;
        addResult(category, "Manual/Enhanced Bidding Usage", "Warn", "Campaign '" + campaignName + "' uses " + biddingStrategyType + ".", "Ensure manual/eCPC bidding is intentional and actively managed. Consider Smart Bidding if sufficient conversion data exists.");
      } else if (["TARGET_CPA", "TARGET_ROAS", "MAXIMIZE_CONVERSIONS", "MAXIMIZE_CONVERSION_VALUE"].indexOf(biddingStrategyType) !== -1) {
        smartBiddingCampaigns++;
         addResult(category, "Smart Bidding Usage", "Pass", "Campaign '" + campaignName + "' uses " + biddingStrategyType + ".", "Ensure sufficient conversion data (~15-30 conversions in last 30 days recommended) for optimal performance. Monitor targets (tCPA/tROAS).");
      } else if (biddingStrategyType !== "UNKNOWN") {
        otherBiddingCampaigns++;
         addResult(category, "Other Bidding Strategy", "Info", "Campaign '" + campaignName + "' uses " + biddingStrategyType + ".", "Ensure strategy aligns with the primary goal of the campaign (e.g., Max Clicks for traffic, Target IS for visibility).");
      } else {
           addResult(category, "Bidding Strategy Check", "Warn", "Could not determine bidding strategy for campaign '" + campaignName + "'.", "Review manually.");
      }

      // Budget Constraints & Impression Share Loss (Only if stats were available)
      if (stats) {
          if (isLostBudget > CONFIG.MAX_IMPRESSION_SHARE_LOST_BUDGET) {
            limitedByBudgetCampaigns++;
            highIsLostBudgetCampaigns++;
            addResult(category, "Impression Share Lost (Budget)", "Fail", "Campaign '" + campaignName + "' lost " + (isLostBudget * 100).toFixed(1) + "% IS due to budget (>" + (CONFIG.MAX_IMPRESSION_SHARE_LOST_BUDGET * 100) + "%).", "Increase budget if performance is good, or optimize bids/targeting to reduce costs.");
          }
          if (isLostRank > CONFIG.MAX_IMPRESSION_SHARE_LOST_RANK) {
             highIsLostRankCampaigns++;
             addResult(category, "Impression Share Lost (Rank)", "Fail", "Campaign '" + campaignName + "' lost " + (isLostRank * 100).toFixed(1) + "% IS due to rank (>" + (CONFIG.MAX_IMPRESSION_SHARE_LOST_RANK * 100) + "%).", "Improve Quality Score (ad relevance, CTR, landing page) and/or increase bids.");
          }
      }

      // Bid Adjustments (Requires checking targeting criteria)
      var bidAdjustmentCount = 0;
      try {
          // Example: Check device bid adjustments
          var platforms = campaign.targeting().platforms().get();
          while(platforms.hasNext()){
              var platform = platforms.next();
              if(platform.getBidModifier() !== 1.0) bidAdjustmentCount++; // Modifier exists if not 1.0
          }
          // Add checks for Location, Ad Schedule, Audience adjustments similarly
           addResult(category, "Bid Adjustments", bidAdjustmentCount > 0 ? "Info" : "Warn", "Campaign '" + campaignName + "' has " + bidAdjustmentCount + " device bid adjustment(s) checked.", bidAdjustmentCount > 0 ? "Review performance data to ensure adjustments are justified." : "Consider setting bid adjustments for devices, locations, times, or audiences based on performance differences.");
      } catch (adjError) {
           addResult(category, "Bid Adjustments Check", "Error", "Could not check bid adjustments for campaign '" + campaignName + "': " + adjError, "Check permissions/campaign type.");
      }


    }

    // Summary Results
    addResult(category, "Bidding Strategy Mix", "Info", "Manual/eCPC: " + manualCpcCampaigns + ", Smart Bidding: " + smartBiddingCampaigns + ", Other: " + otherBiddingCampaigns, "Ensure the mix of strategies aligns with overall account goals.");
    if (campaignsChecked > 0) {
        if (statsErrorCount > 0) {
             addResult(category, "Impression Share Checks", "Warn", "IS Lost (Budget/Rank) checks skipped for " + statsErrorCount + " campaigns due to stats errors/unavailability.", "Review IS manually for these campaigns.");
        } else {
            if (highIsLostBudgetCampaigns === 0) addResult(category, "Impression Share Lost (Budget)", "Pass", "No campaigns found losing significant IS (>"+(CONFIG.MAX_IMPRESSION_SHARE_LOST_BUDGET * 100)+"%) to budget.", "Monitor budget utilization.");
            else addResult(category, "Budget Constraints", "Fail", limitedByBudgetCampaigns + " campaign(s) potentially limited by budget.", "Review campaigns flagged for high IS Lost (Budget).");

            if (highIsLostRankCampaigns === 0) addResult(category, "Impression Share Lost (Rank)", "Pass", "No campaigns found losing significant IS (>"+(CONFIG.MAX_IMPRESSION_SHARE_LOST_RANK * 100)+"%) to rank.", "Maintain good Quality Scores and competitive bids.");
             else addResult(category, "Rank Constraints", "Fail", highIsLostRankCampaigns + " campaign(s) potentially limited by rank.", "Review campaigns flagged for high IS Lost (Rank)."); // Added else for rank fail summary
        }
    } else {
         addResult(category, "General Check", "Info", "No enabled campaigns found to check bidding.", "N/A");
    }
    addResult(category, "Smart Bidding Optimization", "Info", "Manual Review Required", "For Smart Bidding campaigns, ensure conversion tracking is accurate and review performance against tCPA/tROAS targets. Check Recommendations tab for optimization suggestions.");


  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits Quality Score distribution.
 * Checklist Items: QS review, Low QS (< threshold), Ad relevance (component),
 *                  Expected CTR (component), Landing page experience (component).
 */
function auditQualityScore() {
  var category = "Quality Score";
  Logger.log("--- Auditing " + category + " ---");
  var keywordsChecked = 0; // Keywords attempted
  var lowQsKeywords = 0;
  var keywordsWithQs = 0; // Keywords with valid QS data
  var avgQsSum = 0;
  var lowAdRelevance = 0;
  var lowExpCtr = 0;
  var lowLandingPage = 0;
  var qsErrorCount = 0;

  try {
    var keywordIterator = AdsApp.keywords()
      .withCondition("Status = ENABLED")
      .withCondition("AdGroupStatus = ENABLED")
      .withCondition("CampaignStatus = ENABLED")
      // .withCondition("QualityScore > 0") // Check all enabled, handle null QS below
      .get();

    while (keywordIterator.hasNext()) {
      var keyword = keywordIterator.next();
      keywordsChecked++;
      var qs = null;
      var adGroupName = keyword.getAdGroup().getName();
      var campaignName = keyword.getCampaign().getName();
      var keywordText = keyword.getText();
      var matchType = keyword.getMatchType();

      // Check if QS is valid number
      try {
          qs = keyword.getQualityScore();
          if (qs !== null && !isNaN(qs) && qs > 0) {
              keywordsWithQs++;
              avgQsSum += qs;

              // Low QS Check
              if (qs < CONFIG.MIN_QUALITY_SCORE) {
                lowQsKeywords++;
                var components = [];
                var adRelevance = "N/A";
                var expCtr = "N/A";
                var landingPageExp = "N/A";

                // Check components (use try-catch as they might be null/unavailable)
                try { adRelevance = keyword.getAdRelevance() || "N/A"; } catch(e) {}
                try { expCtr = keyword.getExpectedCtr() || "N/A"; } catch(e) {}
                try { landingPageExp = keyword.getLandingPageExperience() || "N/A"; } catch(e) {}


                if (adRelevance === "BELOW_AVERAGE") { lowAdRelevance++; components.push("Ad Relevance"); }
                if (expCtr === "BELOW_AVERAGE") { lowExpCtr++; components.push("Exp. CTR"); }
                if (landingPageExp === "BELOW_AVERAGE") { lowLandingPage++; components.push("Landing Page Exp."); }

                addResult(category, "Low Quality Score (<" + CONFIG.MIN_QUALITY_SCORE + ")", "Fail", "Keyword '" + keywordText + "' (" + matchType + ") in Ad Group '" + adGroupName + "' has QS: " + qs + ". Below Average Components: [" + components.join(', ') + "]", "Improve the flagged components: tighten ad group themes, improve ad copy, check landing page relevance/speed.");
              }
          } else {
              // Log if QS is null/invalid for a keyword expected to have one (e.g. has impressions)
              // Logger.log("Info: Keyword '" + keywordText + "' in Ad Group '" + adGroupName + "' has no QS data (Value: " + qs + "). Might have low impressions.");
          }
      } catch (e) {
          qsErrorCount++;
          Logger.log("Warning: Could not get QS for keyword '" + keywordText + "' in Ad Group '" + adGroupName + "'. Error: " + e);
      }
    }

    // Summary Results
    if (keywordsWithQs > 0) {
        var avgQs = (avgQsSum / keywordsWithQs).toFixed(1);
        addResult(category, "Average Quality Score", "Info", "Avg. QS for keywords with score: " + avgQs + " (based on " + keywordsWithQs + " keywords).", "Aim to improve overall QS. Benchmark against industry standards if possible.");
        if (lowQsKeywords === 0) {
            addResult(category, "Low Quality Scores (<" + CONFIG.MIN_QUALITY_SCORE + ")", "Pass", "No keywords found with QS below " + CONFIG.MIN_QUALITY_SCORE + ".", "Maintain high relevance across keywords, ads, and landing pages.");
        } else {
             addResult(category, "Low Quality Score Summary", "Fail", lowQsKeywords + " keywords found with QS < " + CONFIG.MIN_QUALITY_SCORE + ". Low Components: Ad Relevance (" + lowAdRelevance + "), Exp. CTR (" + lowExpCtr + "), Landing Page (" + lowLandingPage + ").", "Focus improvement efforts on the most common low components identified in individual keyword results.");
        }
         addResult(category, "Ad Relevance Component", "Info", lowAdRelevance + " keywords flagged with Below Average Ad Relevance.", "Ensure keywords are tightly themed within ad groups and reflected in ad copy.");
         addResult(category, "Expected CTR Component", "Info", lowExpCtr + " keywords flagged with Below Average Expected CTR.", "Improve ad copy visibility, use compelling CTAs, leverage ad extensions, refine keyword targeting.");
         addResult(category, "Landing Page Experience Component", "Info", lowLandingPage + " keywords flagged with Below Average Landing Page Experience.", "Ensure landing pages are relevant to keywords/ads, load quickly, are mobile-friendly, and provide a good user experience.");

    } else if (keywordsChecked > 0) {
         addResult(category, "General Check", "Info", "No keywords with Quality Score data found among " + keywordsChecked + " checked.", "Ensure campaigns are running and keywords have enough impressions to generate QS data.");
    } else {
         addResult(category, "General Check", "Info", "No enabled keywords found to check Quality Score.", "N/A");
    }
    if (qsErrorCount > 0) {
         addResult(category, "Quality Score Errors", "Warn", "Could not retrieve QS for " + qsErrorCount + " keywords.", "Check script logs for details. May indicate permission issues or API changes.");
    }

  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits landing pages: checks URL validity (HTTPS), mobile-friendliness (basic). Speed/content are manual.
 * Checklist Items: Alignment (manual), Headlines/CTAs (manual), Load speed (manual),
 *                  Secure URLs (HTTPS), Broken links (basic check), Keyword incorporation (manual), Mobile-friendly.
 */
function auditLandingPages() {
  var category = "Landing Pages";
  Logger.log("--- Auditing " + category + " ---");
  var urlsCheckedCount = 0;
  var httpUrls = 0;
  var potentiallyBrokenUrls = 0; // Based on simple fetch status
  var nonMobileFriendlyUrls = 0; // Based on AdsApp check
  var checkedUrls = {}; // Avoid re-checking the same URL {url: true}

  try {
      // Ensure UrlFetchApp is available
      if (typeof UrlFetchApp === 'undefined') {
          addResult(category, "URL Fetch Check", "Error", "UrlFetchApp service is not available.", "Cannot perform broken link checks.");
          // Continue with other checks if possible, or return
      }

      var urlSources = [];

      // Get URLs from Enabled Ads (Responsive Search, Expanded Text)
      var adIterator = AdsApp.ads()
          .withCondition("Status = ENABLED")
          .withCondition("AdGroupStatus = ENABLED")
          .withCondition("CampaignStatus = ENABLED")
          .withCondition("Type IN [RESPONSIVE_SEARCH_AD, EXPANDED_TEXT_AD]")
          .get();
      while (adIterator.hasNext()) {
          var ad = adIterator.next();
          try {
              var finalUrl = ad.urls().getFinalUrl();
              if (finalUrl) urlSources.push({ url: finalUrl, context: "Ad in AG: " + ad.getAdGroup().getName() });
          } catch (urlError) { Logger.log("Warning: Could not get Final URL for ad ID " + ad.getId() + ". " + urlError); }
      }

      // Get URLs from Enabled Keywords
      var keywordIterator = AdsApp.keywords()
          .withCondition("Status = ENABLED")
          .withCondition("AdGroupStatus = ENABLED")
          .withCondition("CampaignStatus = ENABLED")
          // .withCondition("FinalUrl != ''") // Removed invalid condition
          .get();
       while (keywordIterator.hasNext()) {
          var keyword = keywordIterator.next();
          try {
              var kwFinalUrl = keyword.urls().getFinalUrl();
               if (kwFinalUrl) urlSources.push({ url: kwFinalUrl, context: "Keyword '" + keyword.getText() + "' in AG: " + keyword.getAdGroup().getName() });
          } catch (urlError) { Logger.log("Warning: Could not get Final URL for keyword ID " + keyword.getId() + ". " + urlError); }
      }

      // Get URLs from Enabled Sitelink Extensions (Older API)
      // Note: Accessing individual final URLs for sitelinks is complex here.
      // We'll just note that sitelinks exist and recommend manual URL checks.
      addResult(category, "Sitelink URL Check", "Info", "Manual check recommended for Sitelink URLs.", "Verify links within Sitelink extensions in the UI.");


      // Process unique URLs up to the sample size limit
      for (var i = 0; i < urlSources.length && urlsCheckedCount < CONFIG.LANDING_PAGE_SAMPLE_SIZE; i++) {
          var source = urlSources[i];
          var url = source.url;
          var context = source.context;

          if (url && !checkedUrls[url]) {
              checkedUrls[url] = true; // Mark as checked
              urlsCheckedCount++;

              // 1. Check HTTPS
              if (!url.toLowerCase().startsWith("https://")) {
                  httpUrls++;
                  addResult(category, "Secure URLs (HTTPS)", "Fail", "URL '" + url + "' (Context: " + context + ") is not HTTPS.", "Update landing page URL to use HTTPS for security and user trust.");
              }

              // 2. Basic Broken Link Check (using UrlFetchApp)
              if (typeof UrlFetchApp !== 'undefined') {
                  try {
                      var options = {
                          'muteHttpExceptions': true,
                          'validateHttpsCertificates': false, // Be lenient
                          'followRedirects': true,
                          'method' : 'GET' // Explicitly set method
                      };
                      var response = UrlFetchApp.fetch(url, options);
                      var responseCode = response.getResponseCode();

                      if (responseCode >= 400) { // 4xx or 5xx errors
                          potentiallyBrokenUrls++;
                          addResult(category, "Potential Broken Link", "Fail", "URL '" + url + "' (Context: " + context + ") returned HTTP status code: " + responseCode, "Verify the page loads correctly. Check for typos or server issues.");
                      }
                  } catch (fetchError) {
                      potentiallyBrokenUrls++; // Count fetch errors as potential issues
                      addResult(category, "URL Fetch Error", "Warn", "Could not fetch URL '" + url + "' (Context: " + context + "): " + fetchError, "Verify the page loads correctly. Could be a temporary issue, redirect loop, or script limitation.");
                  }
              }

              // 3. Mobile Friendliness (API check is often unreliable/deprecated)
              // Rely on manual check or external tools.
              // try {
              //     // Find an ad associated with this URL to check mobile friendly status (complex mapping)
              //     // var mobileFriendly = ad.isMobileFriendly(); // Needs the specific ad object
              //     // if (mobileFriendly === false) { ... }
              // } catch (mfError) { ... }


              // Prevent script timeout by sleeping occasionally
              if (urlsCheckedCount % 20 === 0) {
                  Utilities.sleep(1000);
              }
          } // End if URL not checked
      } // End URL processing loop


    // Summary Results
    if (urlsCheckedCount > 0) {
        if (httpUrls === 0) addResult(category, "Secure URLs (HTTPS)", "Pass", "All " + urlsCheckedCount + " checked unique URLs use HTTPS.", "Maintain HTTPS for all landing pages.");
        if (potentiallyBrokenUrls === 0) addResult(category, "Broken Links Check", "Pass", "No potentially broken links found among " + urlsCheckedCount + " checked unique URLs (based on HTTP status).", "Continue monitoring links.");
        // Mobile friendly summary removed due to API unreliability
    } else {
      addResult(category, "General Check", "Info", "No unique Final URLs found in checked ads/keywords or checking was disabled/limited.", "Ensure ads/keywords have valid final URLs.");
    }

    // Manual Checks
    addResult(category, "Mobile-Friendly Check", "Info", "Manual Check Required", "Use Google's Mobile-Friendly Test tool for reliable checks. Ensure pages are responsive.");
    addResult(category, "Landing Page Alignment", "Info", "Manual Review Required", "Ensure landing page content is highly relevant to the ad copy and keywords that trigger the ad.");
    addResult(category, "Headlines & CTAs on Page", "Info", "Manual Review Required", "Verify landing pages have clear headlines, compelling value propositions, and strong calls-to-action.");
    addResult(category, "Page Load Speed", "Info", "Manual Check Required", "Use Google PageSpeed Insights or similar tools to test landing page load times on desktop and mobile. Optimize images, code, and server response time.");
    addResult(category, "Keyword Incorporation", "Info", "Manual Review Required", "Check if relevant keywords are naturally incorporated into landing page headlines and copy.");


  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred during landing page checks: " + e, "Investigate the error. Consider reducing LANDING_PAGE_SAMPLE_SIZE or disabling checks if timeouts occur.");
  }
}


/**
 * Audits audience targeting usage and exclusions. Performance analysis is manual/metric-based.
 * Checklist Items: Segment definition (usage check), Demographic targeting, Exclusions,
 *                  Performance analysis (manual), Remarketing lists active.
 */
function auditAudienceTargeting() {
  var category = "Audience Targeting";
  Logger.log("--- Auditing " + category + " ---");
  var campaignsChecked = 0;
  var adGroupsChecked = 0;
  var audiencesTargetedCampaign = 0;
  var audiencesTargetedAdGroup = 0;
  var audienceExclusionsCampaign = 0;
  var audienceExclusionsAdGroup = 0;
  var demographicTargetsCampaign = 0; // Age, Gender
  var demographicExclusionsCampaign = 0;
  var remarketingListsUsed = 0; // Count usage instances
  var activeUserLists = 0; // Count available lists

  try {
      // Check User Lists (Remarketing, Custom Audiences) Availability - Manual Check
      addResult(category, "Remarketing/User Lists Available", "Info", "Manual Check Required", "Check available User Lists (Audiences) in the UI (Tools & Settings > Shared Library > Audience manager).");
      // Cannot reliably get activeUserLists count via standard script.


    var campaignIterator = AdsApp.campaigns()
      .withCondition("Status = ENABLED")
      .get();

    while (campaignIterator.hasNext()) {
      var campaign = campaignIterator.next();
      campaignsChecked++;
      var campaignName = campaign.getName();
      var campaignId = campaign.getId(); // For context if needed
      var campaignTargeting = campaign.targeting();
      var campaignHasAudienceTarget = false;
      var campaignHasAudienceExclusion = false;
      var campaignHasDemoTarget = false;
      var campaignHasDemoExclusion = false;


      // Campaign Level Audiences (Targeting & Exclusions)
      try {
          var campAudiences = campaignTargeting.audiences().get();
          while(campAudiences.hasNext()){
              var aud = campAudiences.next();
              audiencesTargetedCampaign++;
              campaignHasAudienceTarget = true;
              if (aud.getUserList() != null) remarketingListsUsed++;
          }
          var campExcludedAudiences = campaignTargeting.excludedAudiences().get();
          while(campExcludedAudiences.hasNext()){
              audienceExclusionsCampaign++;
              campaignHasAudienceExclusion = true;
              campExcludedAudiences.next();
          }
      } catch (audError) {
           addResult(category, "Campaign Audience Check", "Error", "Could not check audiences for campaign '" + campaignName + "': " + audError, "Check permissions/campaign type.");
      }


      // Campaign Level Demographics (Age, Gender) - Manual Check
      addResult(category, "Campaign Demographics Check", "Info", "Manual Check Required for Campaign '" + campaignName + "'", "Review demographic targeting (Age, Gender) in the campaign's Audience settings in the UI.");
      // Commenting out unavailable methods:
      // try {
      //     var genders = campaignTargeting.genders().get(); while(genders.hasNext()){ demographicTargetsCampaign++; campaignHasDemoTarget = true; genders.next();}
      //     var ages = campaignTargeting.ages().get(); while(ages.hasNext()){ demographicTargetsCampaign++; campaignHasDemoTarget = true; ages.next();}
      //     var excludedGenders = campaignTargeting.excludedGenders().get(); while(excludedGenders.hasNext()){ demographicExclusionsCampaign++; campaignHasDemoExclusion = true; excludedGenders.next();}
      //     var excludedAges = campaignTargeting.excludedAges().get(); while(excludedAges.hasNext()){ demographicExclusionsCampaign++; campaignHasDemoExclusion = true; excludedAges.next();}
      // } catch (demoError) {
      //      addResult(category, "Campaign Demographics Check", "Error", "Could not check demographics for campaign '" + campaignName + "': " + demoError, "Check permissions/campaign type.");
      // }


      var adGroupIterator = campaign.adGroups().withCondition("Status = ENABLED").get();
      while (adGroupIterator.hasNext()) {
        var adGroup = adGroupIterator.next();
        adGroupsChecked++;
        var adGroupTargeting = adGroup.targeting();
        var agHasAudienceTarget = false;
        var agHasAudienceExclusion = false;

        // Ad Group Level Audiences
        try {
             var agAudiences = adGroupTargeting.audiences().get();
             while(agAudiences.hasNext()){
                 var agAud = agAudiences.next();
                 audiencesTargetedAdGroup++;
                 agHasAudienceTarget = true; // Mark AG level
                 campaignHasAudienceTarget = true; // Also mark campaign level
                 if (agAud.getUserList() != null) remarketingListsUsed++;
             }
             var agExcludedAudiences = adGroupTargeting.excludedAudiences().get();
             while(agExcludedAudiences.hasNext()){
                 audienceExclusionsAdGroup++;
                 agHasAudienceExclusion = true; // Mark AG level
                 campaignHasAudienceExclusion = true; // Also mark campaign level
                 agExcludedAudiences.next();
             }
        } catch (agAudError) {
             addResult(category, "Ad Group Audience Check", "Error", "Could not check audiences for ad group '" + adGroup.getName() + "': " + agAudError, "Check permissions.");
        }

         // Log ad group specific warnings if needed (e.g., if AG has no targeting but campaign does)
         // if (!agHasAudienceTarget && campaignHasAudienceTarget) { ... }
      } // End ad group loop

       // Campaign level summary warnings
       if (!campaignHasAudienceTarget && !campaignHasDemoTarget) {
            addResult(category, "Audience Targeting Usage", "Warn", "Campaign '" + campaignName + "' does not appear to use audience or demographic targeting at campaign or ad group level.", "Consider adding relevant audiences or demographic refinements (especially for Display/Video or Observation for Search).");
       }
       if (!campaignHasAudienceExclusion && !campaignHasDemoExclusion) {
            addResult(category, "Audience Exclusions", "Warn", "Campaign '" + campaignName + "' does not appear to use audience or demographic exclusions at campaign or ad group level.", "Consider excluding irrelevant audiences or demographics.");
       }


    } // End campaign loop

    // Overall Summary Results
    addResult(category, "Audience Segments Targeted", "Info", (audiencesTargetedCampaign + audiencesTargetedAdGroup) + " audience criteria instances found at campaign/ad group level.", "Ensure targeted audiences align with campaign goals.");
    addResult(category, "Demographic Targeting", "Info", "Manual Check Required", "Review demographic targeting settings in the UI.");
    addResult(category, "Audience/Demographic Exclusions", "Info", (audienceExclusionsCampaign + audienceExclusionsAdGroup) + " audience exclusion criteria instances found.", "Verify audience exclusions are correctly removing irrelevant traffic/users. Review demographic exclusions manually.");
    addResult(category, "Remarketing Lists Used", remarketingListsUsed > 0 ? "Pass" : "Warn", remarketingListsUsed + " instances of user lists (remarketing/custom) used in targeting.", remarketingListsUsed > 0 ? "Ensure lists are sufficiently large and targeted appropriately." : "Consider implementing or expanding remarketing efforts.");

    // Manual Checks
    addResult(category, "Audience Performance Analysis", "Info", "Manual Review Required", "Analyze performance reports segmented by audience (UI: Audiences section). Optimize bids or refine targeting based on performance (CTR, Conv. Rate, CPA).");


  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits key performance metrics against configured thresholds.
 * Checklist Items: CTR, CPC, Conversion Rate, CPA, ROAS, Impression Share, IS Lost (Rank/Budget).
 */
function auditPerformanceMetrics() {
  var category = "Performance Metrics";
  Logger.log("--- Auditing " + category + " (Account Level - Last 30 Days) ---");
  var stats = null;

  try {
    stats = AdsApp.currentAccount().getStatsFor("LAST_30_DAYS");
  } catch (e) {
      addResult(category, "General Check", "Error", "An error occurred fetching account stats: " + e, "Could not retrieve account performance metrics.");
      return; // Exit if basic stats retrieval fails
  }

  // Initialize metrics
  var ctr = 0, avgCpc = 0, convRate = 0, conversions = 0, cost = 0, cpa = 0, convValue = 0, roas = 0, imprShare = 0, isLostRank = 0, isLostBudget = 0;
  var metricErrors = []; // Store names of metrics that failed

  // Try getting each metric individually
  try { ctr = stats.getCtr(); } catch (e) { metricErrors.push("CTR"); }
  try { avgCpc = stats.getAverageCpc(); } catch (e) { metricErrors.push("AvgCPC"); }
  try { convRate = stats.getConversionRate(); } catch (e) { metricErrors.push("ConvRate"); }
  try { conversions = stats.getConversions(); } catch (e) { metricErrors.push("Conversions"); }
  try { cost = stats.getCost(); } catch (e) { metricErrors.push("Cost"); }
  // try { convValue = stats.getConversionValue(); } catch (e) { metricErrors.push("ConvValue"); } // Method likely doesn't exist
  try { imprShare = stats.getSearchImpressionShare(); } catch (e) { metricErrors.push("ImprShare"); }
  try { isLostRank = stats.getSearchRankLostImpressionShare(); } catch (e) { metricErrors.push("ISLostRank"); }
  try { isLostBudget = stats.getSearchBudgetLostImpressionShare(); } catch (e) { metricErrors.push("ISLostBudget"); }

  // Report errors if any metrics failed
  if (metricErrors.length > 0) {
      addResult(category, "Metric Retrieval", "Warn", "Could not retrieve some account metrics: " + metricErrors.join(', '), "Checks involving these metrics may be incomplete. Review manually or check API documentation.");
  }

  // Calculate derived metrics if possible
  if (conversions > 0 && cost !== null) { // Check cost is not null before division
      cpa = cost / conversions;
  }
  // ROAS calculation removed as getConversionValue() is unavailable

  // Perform checks based on available metrics

  // CTR
  if (metricErrors.indexOf("CTR") === -1) {
      if (ctr >= CONFIG.MIN_CTR) {
        addResult(category, "Click-Through Rate (CTR)", "Pass", "Account CTR: " + (ctr * 100).toFixed(2) + "% (>= " + (CONFIG.MIN_CTR * 100) + "%)", "CTR meets minimum threshold. Compare to industry benchmarks.");
      } else {
        addResult(category, "Click-Through Rate (CTR)", "Fail", "Account CTR: " + (ctr * 100).toFixed(2) + "% (< " + (CONFIG.MIN_CTR * 100) + "%)", "Investigate low CTR: improve ad copy, keyword relevance, targeting, or use negative keywords.");
      }
  } else {
       addResult(category, "Click-Through Rate (CTR)", "Info", "CTR check skipped due to retrieval error.", "Review manually.");
  }

  // CPC (Info - highly variable)
  if (metricErrors.indexOf("AvgCPC") === -1) {
      addResult(category, "Average Cost-Per-Click (CPC)", "Info", "Account Avg. CPC: " + avgCpc.toFixed(2), "Monitor CPC trends. Compare to keyword/industry benchmarks. High CPC may impact CPA.");
  } else {
       addResult(category, "Average Cost-Per-Click (CPC)", "Info", "Avg. CPC check skipped due to retrieval error.", "Review manually.");
  }

  // Conversion Rate
  if (metricErrors.indexOf("ConvRate") === -1) {
      if (convRate >= CONFIG.MIN_CONVERSION_RATE) {
        addResult(category, "Conversion Rate", "Pass", "Account Conv. Rate: " + (convRate * 100).toFixed(2) + "% (>= " + (CONFIG.MIN_CONVERSION_RATE * 100) + "%)", "Conversion rate meets minimum threshold. Focus on landing page optimization and offer relevance.");
      } else {
        addResult(category, "Conversion Rate", "Fail", "Account Conv. Rate: " + (convRate * 100).toFixed(2) + "% (< " + (CONFIG.MIN_CONVERSION_RATE * 100) + "%)", "Investigate low conversion rate: check landing page experience, offer clarity, audience targeting, and conversion tracking setup.");
      }
  } else {
       addResult(category, "Conversion Rate", "Info", "Conv. Rate check skipped due to retrieval error.", "Review manually.");
  }

  // CPA
  if (metricErrors.indexOf("Conversions") === -1 && metricErrors.indexOf("Cost") === -1) {
      if (conversions > 0) { // Only evaluate CPA if conversions exist
          if (cpa <= CONFIG.MAX_CPA) {
            addResult(category, "Cost Per Acquisition (CPA)", "Pass", "Account CPA: " + cpa.toFixed(2) + " (<= " + CONFIG.MAX_CPA.toFixed(2) + ")", "CPA is within the acceptable threshold.");
          } else {
            addResult(category, "Cost Per Acquisition (CPA)", "Fail", "Account CPA: " + cpa.toFixed(2) + " (> " + CONFIG.MAX_CPA.toFixed(2) + ")", "Investigate high CPA: optimize bids, improve Quality Score, refine targeting, enhance conversion rates, or review CPA goal.");
          }
      } else {
           addResult(category, "Cost Per Acquisition (CPA)", "Info", "No conversions recorded in the last 30 days. CPA cannot be calculated.", "Ensure conversion tracking is working correctly.");
      }
  } else {
       addResult(category, "Cost Per Acquisition (CPA)", "Info", "CPA check skipped due to retrieval error for Cost or Conversions.", "Review manually.");
  }


  // ROAS (Info - requires value tracking) - Marked as unavailable
   addResult(category, "Return On Ad Spend (ROAS)", "Info", "ROAS check skipped (getConversionValue unavailable).", "Implement conversion value tracking and review ROAS manually if applicable.");


  // Impression Share (Check if Search IS is available)
  if (metricErrors.indexOf("ImprShare") === -1 || metricErrors.indexOf("ISLostRank") === -1 || metricErrors.indexOf("ISLostBudget") === -1) {
      if (imprShare > 0 || isLostRank > 0 || isLostBudget > 0) { // If any search IS metric was retrieved
          if (metricErrors.indexOf("ImprShare") === -1) {
              if (imprShare >= CONFIG.MIN_IMPRESSION_SHARE) {
                addResult(category, "Search Impression Share", "Pass", "Account Search IS: " + (imprShare * 100).toFixed(1) + "% (>= " + (CONFIG.MIN_IMPRESSION_SHARE * 100) + "%)", "Impression share meets minimum threshold.");
              } else {
                 addResult(category, "Search Impression Share", "Warn", "Account Search IS: " + (imprShare * 100).toFixed(1) + "% (< " + (CONFIG.MIN_IMPRESSION_SHARE * 100) + "%)", "Investigate reasons for low impression share (budget or rank).");
              }
          } else {
               addResult(category, "Search Impression Share", "Info", "Search IS check skipped due to retrieval error.", "Review manually.");
          }

          // IS Lost Rank
          if (metricErrors.indexOf("ISLostRank") === -1) {
              if (isLostRank <= CONFIG.MAX_IMPRESSION_SHARE_LOST_RANK) {
                addResult(category, "Search IS Lost (Rank)", "Pass", "Account IS Lost (Rank): " + (isLostRank * 100).toFixed(1) + "% (<= " + (CONFIG.MAX_IMPRESSION_SHARE_LOST_RANK * 100) + "%)", "Impression share loss due to rank is within acceptable limits.");
              } else {
                addResult(category, "Search IS Lost (Rank)", "Fail", "Account IS Lost (Rank): " + (isLostRank * 100).toFixed(1) + "% (> " + (CONFIG.MAX_IMPRESSION_SHARE_LOST_RANK * 100) + "%)", "Improve Quality Score and/or increase bids to regain impression share lost to rank.");
              }
          } else {
               addResult(category, "Search IS Lost (Rank)", "Info", "Search IS Lost (Rank) check skipped due to retrieval error.", "Review manually.");
          }

          // IS Lost Budget
          if (metricErrors.indexOf("ISLostBudget") === -1) {
              if (isLostBudget <= CONFIG.MAX_IMPRESSION_SHARE_LOST_BUDGET) {
                addResult(category, "Search IS Lost (Budget)", "Pass", "Account IS Lost (Budget): " + (isLostBudget * 100).toFixed(1) + "% (<= " + (CONFIG.MAX_IMPRESSION_SHARE_LOST_BUDGET * 100) + "%)", "Impression share loss due to budget is within acceptable limits.");
              } else {
                addResult(category, "Search IS Lost (Budget)", "Fail", "Account IS Lost (Budget): " + (isLostBudget * 100).toFixed(1) + "% (> " + (CONFIG.MAX_IMPRESSION_SHARE_LOST_BUDGET * 100) + "%)", "Increase budgets on performing campaigns or optimize efficiency to reduce impression share lost to budget.");
              }
          } else {
               addResult(category, "Search IS Lost (Budget)", "Info", "Search IS Lost (Budget) check skipped due to retrieval error.", "Review manually.");
          }
      } else if (metricErrors.length === 0) { // Only log this if no *other* errors occurred
           addResult(category, "Search Impression Share", "Info", "Search Impression Share data not available (e.g., Display/Video only account).", "N/A for this check.");
      }
  } else {
       addResult(category, "Search Impression Share", "Info", "Search IS checks skipped due to retrieval errors.", "Review manually.");
  }


  // Note: Campaign/Ad Group level checks provide more granular insights.
  addResult(category, "Granular Performance", "Info", "Account-level metrics provide an overview.", "Analyze performance at the campaign, ad group, keyword, and audience levels for specific optimization opportunities.");

} // Closing brace for auditPerformanceMetrics function


/**
 * Audits campaign optimization status based on performance data.
 * Checklist Items: Underperforming campaigns, High-performing campaigns budget,
 *                  Ad schedule optimization, Geo-targeting optimization, Device adjustments.
 */
function auditCampaignOptimization() {
  var category = "Campaign Optimization";
  Logger.log("--- Auditing " + category + " ---");
  var campaignsChecked = 0;
  var statsErrorCount = 0; // Count campaigns where stats failed

  try {
    var campaignIterator = AdsApp.campaigns()
      .withCondition("Status = ENABLED")
      // .forDateRange("LAST_30_DAYS") // Date range not needed if not getting stats
      .get();

    while (campaignIterator.hasNext()) {
      var campaign = campaignIterator.next();
      campaignsChecked++;
      var campaignName = campaign.getName();

      // Log that detailed performance checks are skipped as campaign.getStats() is unavailable
      addResult(category, "Campaign Performance Check", "Info", "Detailed performance checks (CPA, Conv Rate, IS Lost) skipped for campaign '" + campaignName + "' due to API limitations (getStats unavailable).", "Review campaign performance manually using Google Ads reports.");
      statsErrorCount++; // Increment count as we are skipping stats-based checks

      // Ad Schedule, Geo, Device Optimization (Still relevant as manual checks)
       addResult(category, "Ad Schedule Optimization", "Info", "Manual Review Required for Campaign '" + campaignName + "'", "Analyze performance by day/hour in the UI (Reports > Predefined > Time). Apply bid adjustments or ad schedules based on data.");
       addResult(category, "Geo-targeting Optimization", "Info", "Manual Review Required for Campaign '" + campaignName + "'", "Analyze performance by location in the UI (Locations tab). Refine targeting or apply bid adjustments based on data.");
       addResult(category, "Device Optimization", "Info", "Manual Review Required for Campaign '" + campaignName + "'", "Analyze performance by device in the UI (Settings > Devices). Apply bid adjustments based on data.");

    }

    // Summary Results
    if (campaignsChecked > 0) {
        addResult(category, "Underperforming Campaigns", "Info", "Automated check skipped due to API limitations.", "Review campaign performance manually to identify underperformers.");
        addResult(category, "High Performing Campaign Budget", "Info", "Automated check skipped due to API limitations.", "Review performance and budget limitations manually for high-performing campaigns.");
    } else {
      addResult(category, "General Check", "Info", "No enabled campaigns found to check optimization status.", "N/A");
    }
     if (statsErrorCount > 0) {
         // This message is now logged per campaign, so a summary might be redundant.
         // Logger.log("Info: Detailed performance checks were skipped for " + statsErrorCount + " campaigns due to unavailable stats.");
     }

  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits usage of automation features (Rules, Scripts). Recommendations/Tools are manual.
 * Checklist Items: Automated rules, Scripts, Recommendations review, Third-party tools, Automation alignment.
 */
function auditAutomationTools() {
  var category = "Automation & Tools";
  Logger.log("--- Auditing " + category + " ---");

  try {
    // Automated Rules (API Limitation)
    addResult(category, "Automated Rules", "Info", "Manual Check Required", "Review existing automated rules in the UI (Tools & Settings > Bulk Actions > Rules). Ensure they are functioning correctly and align with current goals. Check rule history for errors.");

    // Scripts (API Limitation)
    var scriptId = "N/A";
    try { scriptId = ScriptApp.getScriptId(); } catch(e) {} // Might fail in some contexts
    addResult(category, "Scripts", "Info", "Manual Check Required", "Review other active Google Ads Scripts in the UI (Tools & Settings > Bulk Actions > Scripts). Ensure they are necessary, functioning correctly (check logs), and not conflicting.");
    addResult(category, "This Audit Script", "Info", "This script (ID: " + scriptId + ") is running.", "Schedule this script to run regularly (e.g., weekly/monthly) and monitor its logs.");


    // Google Ads Recommendations (API Limitation)
    addResult(category, "Google Ads Recommendations", "Info", "Manual Check Required", "Regularly review the 'Recommendations' tab in the Google Ads UI. Evaluate each suggestion carefully before applying; not all recommendations are suitable for every account.");

    // Third-Party Tools (API Limitation)
    addResult(category, "Third-Party Tools", "Info", "Manual Check Required", "If using third-party management or reporting tools, ensure they are correctly integrated and providing value. Verify data consistency between tools and Google Ads.");

    // Automation Alignment (Manual Judgement)
    addResult(category, "Automation Alignment", "Info", "Manual Review Required", "Ensure all automation (Rules, Scripts, Smart Bidding, Third-Party Tools) works together cohesively and supports overall campaign objectives. Avoid conflicting automations.");


  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}


/**
 * Audits competitive analysis elements available via API (Auction Insights).
 * Checklist Items: Auction insights review, Competitor positioning (manual), Gaps (manual), Differentiation (manual).
 */
function auditCompetitiveAnalysis() {
  var category = "Competitive Analysis";
  Logger.log("--- Auditing " + category + " ---");

  try {
    // Auction Insights (API Limitation)
    addResult(category, "Auction Insights Report", "Info", "Manual Check Required", "Regularly review the Auction Insights report in the UI (available at Campaign, Ad Group, and Keyword levels). Analyze impression share, overlap rate, position above rate, etc., for key competitors.");
    addResult(category, "Competitor Ad Positioning", "Info", "Manual Review Required", "Use Auction Insights and manual searches to understand how competitors position themselves in their ad copy and what offers they promote.");
    addResult(category, "Identifying Gaps", "Info", "Manual Review Required", "Analyze competitor strategies (keywords, ads, targeting - based on insights and observation) to identify potential gaps or opportunities they might be missing.");
    addResult(category, "Differentiation Opportunities", "Info", "Manual Review Required", "Based on competitor analysis, identify ways to differentiate your ads, offers, or targeting to stand out.");

  } catch (e) {
    addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
  }
}

/**
 * Notes on reporting and insights generation (mostly manual process).
 * Checklist Items: Dashboards, Metric tracking, Trend analysis, Reporting sharing, Actionable insights.
 */
function auditReportingInsights() {
    var category = "Reporting & Insights";
    Logger.log("--- Auditing " + category + " ---");

    try {
        addResult(category, "Custom Dashboards", "Info", "Manual Check Recommended", "Consider creating custom dashboards in Google Ads or Google Data Studio (Looker Studio) to visualize key metrics and trends relevant to your goals.");
        addResult(category, "Key Metric Tracking", "Info", "Script Provides Data", "This audit script provides a snapshot of key metrics. Track these metrics over time using reports or dashboards to monitor progress.");
        addResult(category, "Trend & Anomaly Identification", "Info", "Manual Analysis Required", "Regularly analyze performance data (using reports/dashboards) to identify positive/negative trends or unexpected anomalies that require investigation.");
        addResult(category, "Report Sharing", "Info", "Manual Process", "Share audit summaries (like the generated spreadsheet) and regular performance reports with relevant stakeholders.");
        addResult(category, "Actionable Insights Documentation", "Info", "Manual Process", "Document insights gained from audits and performance analysis, along with the actions taken or planned, to track optimization efforts.");

    } catch (e) {
        addResult(category, "General Check", "Error", "An error occurred: " + e, "Investigate the error.");
    }
}
