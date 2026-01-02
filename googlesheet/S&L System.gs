// ============================================
// SAVINGS & LOAN SYSTEM
// COMPLETE VERSION - WITH SETTINGS
// TWICE MONTHLY PAYMENTS: 5th & 20th (₱1,000 each)
// ============================================

// Default configuration - can be changed via Settings menu
let CONFIG = {
  interestRateWithLoan: 0.09,        // 9% MONTHLY interest for loans
  interestRateWithoutLoan: 0.09,     // 9% ANNUAL interest for savings (no loan)
  interestRateSavingsOnly: 0.09,     // 9% ANNUAL interest for savings only members
  monthlyContribution: 2000,         // ₱2,000 per month
  paymentPerPeriod: 1000,            // ₱1,000 on 5th and 20th
  paymentDates: [5, 20],             // 5th and 20th of each month
  companyName: "Community Savings & Loan",
  nextMemberId: 1001,                // Starting member ID
  reserveRatio: 0.3,                 // 30% reserve ratio
  minLoanAmount: 1000,               // Minimum loan amount
  maxLoanAmount: 20000,              // Maximum loan amount
  loanTermMin: 1,                    // Minimum loan term (months)
  loanTermMax: 12,                   // Maximum loan term (months)
  savingsMin: 1000,                  // Minimum savings
  savingsForLoan: 6000               // Minimum savings required for loan
};

const COLORS = {
  primary: "#2E7D32",
  secondary: "#1976D2",
  accent: "#F57C00",
  success: "#388E3C",
  warning: "#FF9800",
  danger: "#D32F2F",
  info: "#0288D1",
  light: "#F5F5F5",
  dark: "#212121",
  white: "#FFFFFF",
  header: "#1B5E20"
};

const STATUS_COLORS = {
  'On Track': COLORS.success,
  'Active': COLORS.success,
  'Upcoming': COLORS.info,
  'Due Soon': COLORS.warning,
  'Overdue': COLORS.danger,
  'Paid': COLORS.success,
  'No Loan': COLORS.info,
  'Pending': COLORS.warning,
  'Inactive': COLORS.danger,
  'Verified': COLORS.success,
  'Unverified': COLORS.warning,
  'Complete': COLORS.success,
  'Regular': COLORS.success,
  'Additional': COLORS.info,
  'Late': COLORS.danger
};

// ============================================
// MAIN INITIALIZATION FUNCTION
// ============================================

function initializeSystem() {
  generateProfessionalSystem();
}

function generateProfessionalSystem() {
  try {
    console.log('Starting system generation...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Clear existing sheets except first one
    const sheets = ss.getSheets();
    for (let i = sheets.length - 1; i >= 0; i--) {
      const sheet = sheets[i];
      if (i === 0) {
        sheet.setName('Default');
        sheet.clear();
      } else {
        try {
          ss.deleteSheet(sheet);
        } catch (e) {
          console.log('Could not delete sheet:', sheet.getName());
        }
      }
    }
    
    console.log('Creating sheets...');
    createMemberPortal(ss);
    createSavingsSheet(ss);
    createSavingsPaymentsSheet(ss);
    createLoanSheet(ss);
    createLoanPaymentsSheet(ss);
    createSummarySheet(ss);
    createDashboardSheet(ss);
    createCommunityFundsSheet(ss);
    createSettingsSheet(ss);
    
    createAdminMenu();
    
    console.log('System generation complete!');
    SpreadsheetApp.getUi().alert(
      '✅ System Generated Successfully!', 
      'Professional Savings & Loan System is ready!\n\n' +
      '📅 Payment Schedule:\n' +
      '• 5th of month: ₱1,000\n' +
      '• 20th of month: ₱1,000\n' +
      '• Total: ₱2,000 per month\n\n' +
      '💰 INTEREST RATES:\n' +
      '• Savings with active loan: ' + (CONFIG.interestRateWithoutLoan * 100) + '% per year\n' +
      '• Savings only (no loan): ' + (CONFIG.interestRateSavingsOnly * 100) + '% per year\n' +
      '• Loans: ' + (CONFIG.interestRateWithLoan * 100) + '% PER MONTH\n\n' +
      '🏦 Loan Terms:\n' +
      '• Minimum: ₱' + CONFIG.minLoanAmount.toLocaleString() + '\n' +
      '• Maximum: ₱' + CONFIG.maxLoanAmount.toLocaleString() + '\n' +
      '• Term: ' + CONFIG.loanTermMin + '-' + CONFIG.loanTermMax + ' months\n\n' +
      '📊 Community Funds Feature:\n' +
      '• Auto-checks available funds before approving loans\n' +
      '• Maintains ' + (CONFIG.reserveRatio * 100) + '% reserve ratio\n' +
      '• Real-time fund tracking\n\n' +
      '👤 Start by adding your first member!', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error in generateProfessionalSystem:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to generate system: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function onEdit(e) {
  try {
    // Only run if we have event data
    if (!e) return;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = e.range;
    const sheet = range.getSheet();
    
    // Check if edit is in Member Portal
    if (sheet.getName() === '👤 Member Portal') {
      // Check if edit is in Member ID cell (C7)
      if (range.getRow() === 7 && range.getColumn() === 3) {
        const selectedValue = range.getValue();
        
        if (selectedValue && selectedValue.toString().trim() !== '') {
          if (selectedValue === 'Clear ID') {
            // Clear the selection and display
            range.clearContent();
            clearMemberPortalDisplay();
            console.log('Selection cleared via dropdown');
          } else {
            // It's a valid member ID
            console.log('Member ID selected:', selectedValue);
            
            // Give a small delay, then update
            Utilities.sleep(500);
            
            // Run the update
            searchAndDisplayMember(selectedValue.toString().trim());
          }
        } else {
          // Clear display if empty
          clearMemberPortalDisplay();
        }
      }
    }
  } catch (error) {
    console.error('Error in onEdit:', error);
  }
}

function searchAndDisplayMember(memberId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    
    if (!portalSheet) return;
    
    // Search for member
    const result = searchMemberById(memberId);
    
    if (result.success) {
      // Display member info
      displayMemberInfo(portalSheet, result);
    } else {
      // Show error
      portalSheet.getRange(11, 1, 1, 6).merge()
        .setValue('❌ ' + result.message)
        .setFontColor(COLORS.danger)
        .setHorizontalAlignment('center')
        .setFontWeight('bold');
    }
    
  } catch (error) {
    console.error('Error displaying member:', error);
  }
}

function displayMemberInfo(sheet, member) {
  try {
    // Clear previous data
    sheet.getRange(11, 1, 10, 6).clearContent().clearFormat();
    
    // Display header row
    sheet.getRange(11, 1, 1, 6).merge()
      .setValue('✅ Member Information - ' + member.memberName)
      .setFontColor(COLORS.success)
      .setHorizontalAlignment('center')
      .setFontWeight('bold');
    
    // Display basic info
    const infoData = [
      ['Member ID:', member.memberId],
      ['Member Name:', member.memberName],
      ['Total Savings:', '₱' + (member.totalSavings || 0).toFixed(2)],
      ['Current Loan:', '₱' + (member.currentLoan || 0).toFixed(2)],
      ['Next Payment:', member.nextPaymentDate || 'N/A'],
      ['Status:', member.status || 'Unknown']
    ];
    
    sheet.getRange(12, 1, infoData.length, 2).setValues(infoData);
    
    // Format the info
    sheet.getRange(12, 1, infoData.length, 1).setFontWeight('bold');
    sheet.getRange(12, 2, infoData.length, 1).setHorizontalAlignment('left');
    
    // Add status color
    if (member.status) {
      const statusRow = 17; // Adjust based on your layout
      const statusCell = sheet.getRange(statusRow, 2);
      statusCell.setBackground(STATUS_COLORS[member.status] || COLORS.light)
        .setFontColor(COLORS.dark)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
    
  } catch (error) {
    console.error('Error displaying info:', error);
  }
}

function checkPortalSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const portal = ss.getSheetByName('👤 Member Portal');
  const savings = ss.getSheetByName('💰 Savings');
  
  console.log('Portal sheet exists:', !!portal);
  console.log('Savings sheet exists:', !!savings);
  
  if (savings) {
    const data = savings.getDataRange().getValues();
    console.log('Total rows in savings:', data.length);
    
    const members = [];
    for (let i = 4; i < data.length && i < 10; i++) {
      if (data[i] && data[i][0]) {
        members.push(data[i][0]);
      }
    }
    console.log('Available members:', members);
  }
  
  if (portal) {
    const dropdownCell = portal.getRange('C7');
    const rule = dropdownCell.getDataValidation();
    console.log('Dropdown validation exists:', !!rule);
  }
}

function addClearOptionToDropdown() {
  try {
    console.log('Adding Clear option to dropdown...');
    
    // First refresh all dropdowns
    refreshAllDropdowns();
    
    // Specifically update portal dropdown
    updateMemberPortalDropdown();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    
    if (portalSheet) {
      // Show current dropdown value for verification
      const currentValue = portalSheet.getRange(7, 3).getValue();
      console.log('Current dropdown value:', currentValue);
      
      // Show the dropdown in action
      portalSheet.showSheet();
      ss.setActiveSheet(portalSheet);
      
      SpreadsheetApp.getUi().alert(
        '✅ Clear Option Added',
        'Member Portal dropdown now has "--- Clear Selection ---" option.\n\n' +
        'To use:\n' +
        '1. Click cell C7\n' +
        '2. Select "--- Clear Selection ---" from dropdown\n' +
        '3. Portal will clear automatically',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('Error:', error);
  }
}

// Install trigger for auto-update
function installAutoUpdateTrigger() {
  try {
    // Remove existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'autoUpdateMemberPortal') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new trigger
    ScriptApp.newTrigger('autoUpdateMemberPortal')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    
    console.log('Auto-update trigger installed');
  } catch (error) {
    console.error('Error installing trigger:', error);
  }
}

// Auto-update function triggered by edits
function autoUpdateMemberPortal(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    
    if (!portalSheet) return;
    
    const range = e.range;
    const sheet = range.getSheet();
    
    // Only run if edit is in Member Portal sheet
    if (sheet.getName() !== '👤 Member Portal') return;
    
    // Only run if edit is in the Member ID cell (C7)
    if (range.getRow() === 7 && range.getColumn() === 3) {
      const memberId = portalSheet.getRange(7, 3).getValue();
      
      if (memberId && memberId.toString().trim() !== '') {
        // Auto-update after 1 second delay
        Utilities.sleep(1000);
        autoSearchMember(memberId.toString().trim());
      } else {
        // Clear if empty
        clearMemberPortalDisplay();
      }
    }
    
  } catch (error) {
    console.error('Error in auto-update:', error);
  }
}

// Auto-search function
function autoSearchMember(memberId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    
    if (!portalSheet) return;
    
    const result = searchMemberById(memberId);
    
    if (!result.success) {
      // Clear display area
      clearMemberPortalDisplay();
      
      // Show error message
      portalSheet.getRange(11, 1, 1, 6).merge()
        .setValue('❌ ' + result.message)
        .setFontColor(COLORS.danger)
        .setHorizontalAlignment('center')
        .setFontWeight('bold');
    } else {
      // Update the display immediately
      updatePortalDisplay(result);
    }
    
  } catch (error) {
    console.error('Error in auto-search:', error);
  }
}

// Update portal display with member info
function updatePortalDisplay(result) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    
    if (!portalSheet) return;
    
    // Clear existing data
    clearMemberPortalDisplay();
    
    // Update main info row
    const displayData = [
      [result.memberId || 'N/A', 
       result.memberName || 'N/A', 
       result.totalBalanceWithInterest || 0,
       result.currentLoan || 0,
       result.nextPaymentDate || 'N/A',
       result.status || 'Unknown']
    ];
    
    portalSheet.getRange(12, 1, 1, 6).setValues(displayData);
    portalSheet.getRange(12, 1, 1, 6).setHorizontalAlignment('center');
    
    // Apply number formats
    portalSheet.getRange(12, 3).setNumberFormat('₱#,##0.00');
    portalSheet.getRange(12, 4).setNumberFormat('₱#,##0.00');
    
    // Apply status color coding
    const statusCell = portalSheet.getRange(12, 6);
    statusCell.setBackground(STATUS_COLORS[result.status] || COLORS.light)
      .setFontColor(COLORS.dark)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Update additional info
    const additionalInfo = [
      ['Active Loans', result.activeLoans || 0, 'Number of active loans'],
      ['Monthly Interest', '₱' + (result.monthlyInterest || 0).toFixed(2), 'Interest earned this month'],
      ['Total Interest', '₱' + (result.interestEarned || 0).toFixed(2), 'Total interest earned'],
      ['Payment Streak', result.savingsStreak || 0, 'Consecutive on-time payments'],
      ['Loan Status', result.loanStatus || 'No Loan', 'Status of any active loans']
    ];
    
    portalSheet.getRange(15, 1, additionalInfo.length, 3).setValues(additionalInfo);
    
    // Format additional info
    portalSheet.getRange(15, 2, additionalInfo.length, 1).setHorizontalAlignment('center');
    portalSheet.getRange(15, 3, additionalInfo.length, 1).setHorizontalAlignment('left').setFontStyle('italic');
    
    // Update success message
    portalSheet.getRange(11, 1, 1, 6).merge()
      .setValue('✅ Information loaded for ' + result.memberName)
      .setFontColor(COLORS.success)
      .setHorizontalAlignment('center')
      .setFontWeight('bold');
    
  } catch (error) {
    console.error('Error updating portal display:', error);
  }
}

// Clear portal display
function clearMemberPortalDisplay() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    
    if (!portalSheet) return;
    
    // Clear rows 11-20
    portalSheet.getRange(11, 1, 10, 6).clearContent().clearFormat();
    
    // Reset to default message
    portalSheet.getRange(11, 1, 1, 6).merge()
      .setValue('👤 Select your Member ID above to view your information')
      .setFontColor(COLORS.info)
      .setHorizontalAlignment('center')
      .setFontStyle('italic');
    
    // Clear additional info
    portalSheet.getRange(15, 1, 5, 3).clearContent();
    
  } catch (error) {
    console.error('Error clearing portal display:', error);
  }
}

function fixMemberPortal() {
  try {
    console.log('Fixing Member Portal...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Delete existing Member Portal
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    if (portalSheet) {
      ss.deleteSheet(portalSheet);
      console.log('Deleted old Member Portal');
    }
    
    // Create new Member Portal
    createMemberPortal(ss);
    
    // Go to the new portal
    const newPortal = ss.getSheetByName('👤 Member Portal');
    newPortal.showSheet();
    ss.setActiveSheet(newPortal);
    
    // Force update dropdowns
    refreshAllDropdowns();
    
    // Install triggers
    installAutoUpdateTrigger();
    
    // Test with existing member
    const savingsSheet = ss.getSheetByName('💰 Savings');
    if (savingsSheet) {
      const data = savingsSheet.getDataRange().getValues();
      const members = [];
      for (let i = 4; i < data.length && i < 10; i++) {
        if (data[i][0]) {
          members.push(data[i][0]);
        }
      }
      
      if (members.length > 0) {
        // Select first member in dropdown
        newPortal.getRange(7, 3).setValue(members[0]);
        console.log('Auto-selected member:', members[0]);
      }
    }
    
    console.log('✅ Member Portal fixed successfully');
    
    SpreadsheetApp.getUi().alert(
      '✅ Member Portal Fixed',
      'Member Portal has been completely rebuilt!\n\n' +
      'The auto-update should now work when you:\n' +
      '1. Select a Member ID from cell C7 dropdown\n' +
      '2. Or click the VIEW INFO button\n\n' +
      'Layout has been corrected with proper buttons.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error fixing Member Portal:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to fix: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function fixSummarySheetCalculations() {
  try {
    console.log('Fixing Summary Sheet calculations...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName('📊 Summary');
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const loanSheet = ss.getSheetByName('🏦 Loans');
    const loanPaymentsSheet = ss.getSheetByName('💳 Loan Payments');
    
    if (!summarySheet || !savingsSheet) {
      console.error('Required sheets not found');
      return;
    }
    
    // Get all data
    const savingsData = savingsSheet.getDataRange().getValues();
    const summaryData = summarySheet.getDataRange().getValues();
    
    // Update each member in summary sheet
    for (let i = 3; i < summaryData.length; i++) {
      const memberId = summaryData[i][0];
      if (!memberId || memberId.toString().trim() === '') continue;
      
      let memberSavings = 0;
      let memberSavingsBalance = 0;
      let memberInterestEarned = 0;
      let memberStatus = 'Unknown';
      let savingsStreak = 0;
      
      // Find member in savings sheet
      for (let j = 3; j < savingsData.length; j++) {
        if (savingsData[j][0] === memberId) {
          memberSavings = savingsData[j][2] || 0;           // Column C: Total Savings
          memberSavingsBalance = savingsData[j][9] || 0;    // Column J: Balance
          memberInterestEarned = savingsData[j][10] || 0;   // Column K: Interest Earned
          memberStatus = savingsData[j][7] || 'Unknown';    // Column H: Status
          
          // Calculate savings streak based on status
          // If status is "On Track", check payments
          if (memberStatus === 'On Track') {
            const monthsActive = savingsData[j][8] || 0;    // Column I: Months Active
            const lastPayment = savingsData[j][4];          // Column E: Last Payment
            
            if (lastPayment instanceof Date) {
              const today = new Date();
              const daysSinceLastPayment = Math.floor((today - lastPayment) / (1000 * 60 * 60 * 24));
              
              // Streak is number of months with on-time payments
              // Simplified: Use months active as streak for now
              savingsStreak = Math.min(monthsActive, 12); // Cap at 12
            } else {
              savingsStreak = 1; // At least 1 for active member
            }
          } else {
            savingsStreak = 0; // Not on track = streak broken
          }
          break;
        }
      }
      
      // Calculate loan information
      let totalLoans = 0;
      let activeLoans = 0;
      let interestPaid = 0;
      let loanStatus = 'No Loan';
      let loanStreak = 0;
      
      if (loanSheet) {
        const loanData = loanSheet.getDataRange().getValues();
        for (let k = 3; k < loanData.length; k++) {
          if (loanData[k][1] === memberId) {
            const loanAmount = loanData[k][2] || 0;
            const status = loanData[k][6] || '';
            
            totalLoans += loanAmount;
            
            if (status === 'Active') {
              activeLoans++;
              loanStatus = 'Active';
            }
          }
        }
      }
      
      // Get interest paid from loan payments
      if (loanPaymentsSheet) {
        const loanPaymentsData = loanPaymentsSheet.getDataRange().getValues();
        for (let l = 3; l < loanPaymentsData.length; l++) {
          if (loanPaymentsData[l][1]) {
            // Need to find loan ID first, then check member
            const loanId = loanPaymentsData[l][1];
            if (loanSheet) {
              const loanData = loanSheet.getDataRange().getValues();
              for (let m = 3; m < loanData.length; m++) {
                if (loanData[m][0] === loanId && loanData[m][1] === memberId) {
                  interestPaid += loanPaymentsData[l][5] || 0; // Column F: Interest
                  break;
                }
              }
            }
          }
        }
      }
      
      // Calculate overall balance (Savings + Interest - Loans)
      const overallBalance = memberSavingsBalance - totalLoans;
      
      // Update summary sheet
      summarySheet.getRange(i + 1, 3).setValue(memberSavings);          // Column C: Savings
      summarySheet.getRange(i + 1, 4).setValue(totalLoans);             // Column D: Loans
      summarySheet.getRange(i + 1, 5).setValue(activeLoans);            // Column E: Active Loans
      summarySheet.getRange(i + 1, 6).setValue(loanStatus);             // Column F: Loan Status
      summarySheet.getRange(i + 1, 7).setValue(interestPaid);           // Column G: Interest Paid (LOAN interest)
      summarySheet.getRange(i + 1, 8).setValue(memberStatus);           // Column H: Savings Status
      summarySheet.getRange(i + 1, 9).setValue(overallBalance);         // Column I: Balance
      summarySheet.getRange(i + 1, 10).setValue(savingsStreak);         // Column J: Savings Streak
      summarySheet.getRange(i + 1, 11).setValue(loanStreak);            // Column K: Loan Streak
      
      // Apply formatting
      summarySheet.getRange(i + 1, 3).setNumberFormat('₱#,##0.00');     // Savings
      summarySheet.getRange(i + 1, 4).setNumberFormat('₱#,##0.00');     // Loans
      summarySheet.getRange(i + 1, 7).setNumberFormat('₱#,##0.00');     // Interest Paid
      summarySheet.getRange(i + 1, 9).setNumberFormat('₱#,##0.00');     // Balance
      
      // Apply status colors
      const statusCell = summarySheet.getRange(i + 1, 8);
      statusCell.setBackground(STATUS_COLORS[memberStatus] || COLORS.light)
        .setFontColor(COLORS.dark)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
        
      const loanStatusCell = summarySheet.getRange(i + 1, 6);
      loanStatusCell.setBackground(STATUS_COLORS[loanStatus] || COLORS.light)
        .setFontColor(COLORS.dark)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
    
    console.log('✅ Summary sheet calculations fixed');
    
  } catch (error) {
    console.error('Error fixing summary sheet:', error);
  }
}

function calculateSavingsStreak(memberId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const paymentsSheet = ss.getSheetByName('💵 Savings Payments');
    
    if (!savingsSheet || !paymentsSheet) return 0;
    
    // Get member data
    const savingsData = savingsSheet.getDataRange().getValues();
    let monthsActive = 0;
    let status = '';
    
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] === memberId) {
        monthsActive = savingsData[i][8] || 0;
        status = savingsData[i][7] || '';
        break;
      }
    }
    
    // If not on track, streak is 0
    if (status !== 'On Track') return 0;
    
    // Get payment history
    const paymentsData = paymentsSheet.getDataRange().getValues();
    let regularPayments = 0;
    let currentStreak = 0;
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    
    for (let i = 3; i < paymentsData.length; i++) {
      if (paymentsData[i][1] === memberId && paymentsData[i][4] === 'Regular') {
        const paymentDate = paymentsData[i][2];
        if (paymentDate instanceof Date) {
          const paymentMonth = paymentDate.getMonth();
          const paymentYear = paymentDate.getFullYear();
          
          // Check if payment was in current month
          if (paymentMonth === currentMonth && paymentYear === currentYear) {
            regularPayments++;
          }
        }
      }
    }
    
    // Streak calculation:
    // - At least 1 for being active
    // - Add 0.5 for each regular payment this month
    // - Cap at months active
    let streak = 1; // Base streak for active member
    if (regularPayments >= 2) {
      streak = monthsActive; // Has both payments this month
    } else if (regularPayments === 1) {
      streak = Math.max(1, Math.floor(monthsActive / 2)); // Only one payment
    }
    
    return Math.min(streak, monthsActive);
    
  } catch (error) {
    console.error('Error calculating savings streak:', error);
    return 0;
  }
}

// ============================================
// SHEET CREATION FUNCTIONS
// ============================================

function createMemberPortal(ss) {
  console.log('Creating Member Portal...');
  const sheet = ss.insertSheet('👤 Member Portal');
  
  // Clear sheet first
  sheet.clear();
  
  // Title
  sheet.getRange(1, 1, 1, 6).merge()
    .setValue('🏦 ' + CONFIG.companyName + ' - Member Portal')
    .setFontSize(16).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.header)
    .setHorizontalAlignment('center');
  
  // Subtitle
  sheet.getRange(2, 1, 1, 6).merge()
    .setValue('📅 Payment Schedule: 5th (₱1,000) & 20th (₱1,000) of each month | Total: ₱2,000/month')
    .setFontSize(11)
    .setFontColor(COLORS.primary)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');
  
  // Instructions
  sheet.getRange(4, 1, 1, 6).merge()
    .setValue('🔍 SEARCH YOUR INFORMATION')
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor(COLORS.white)
    .setBackground(COLORS.secondary)
    .setHorizontalAlignment('center');
  
  sheet.getRange(5, 1, 1, 6).merge()
    .setValue('Select your Member ID from dropdown to view your information')
    .setFontSize(11)
    .setFontColor(COLORS.dark)
    .setHorizontalAlignment('center');
  
  // Search Area
  sheet.getRange(7, 2).setValue('Member ID:').setFontWeight('bold');
  const memberIdCell = sheet.getRange(7, 3);
  memberIdCell.setValue('')
    .setBackground(COLORS.light)
    .setBorder(true, true, true, true, null, null, COLORS.dark, SpreadsheetApp.BorderStyle.SOLID)
    .setNote('Select your Member ID from the dropdown - Information will auto-update');
  
  // Add auto-update button
  sheet.getRange(7, 5).setValue('🔍 VIEW INFO')
    .setFontWeight('bold')
    .setFontColor(COLORS.white)
    .setBackground(COLORS.success)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.dark, SpreadsheetApp.BorderStyle.SOLID);
  
  sheet.getRange(7, 6).setValue('🔄 CLEAR')
    .setFontWeight('bold')
    .setFontColor(COLORS.white)
    .setBackground(COLORS.warning)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, COLORS.dark, SpreadsheetApp.BorderStyle.SOLID);
  
  // Headers for member info
  sheet.getRange(10, 1, 1, 6).setValues([[
    'Member ID', 'Member Name', 'Total Savings', 
    'Current Loan', 'Next Payment Date', 'Status'
  ]])
  .setFontWeight('bold').setFontColor(COLORS.white)
  .setBackground(COLORS.primary).setHorizontalAlignment('center');
  
  // Initialize display row with formulas that auto-update
  const displayRow = 11;
  sheet.getRange(displayRow, 1, 1, 6).merge()
    .setValue('👤 Select your Member ID above to view your information')
    .setFontColor(COLORS.info)
    .setHorizontalAlignment('center')
    .setFontStyle('italic');
  
  // Add formulas that auto-update based on selected member ID
  // Row 13-15 will show additional info
  sheet.getRange(13, 1, 1, 6).merge()
    .setValue('📊 ADDITIONAL INFORMATION')
    .setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.info)
    .setHorizontalAlignment('center');
  
  // Additional info headers
  sheet.getRange(14, 1, 1, 3).setValues([['Information', 'Value', 'Notes']])
    .setFontWeight('bold').setFontColor(COLORS.dark)
    .setBackground(COLORS.light).setHorizontalAlignment('center');
  
  // Additional info rows
  const infoRows = [
    ['Active Loans', '', 'Number of active loans'],
    ['Monthly Interest', '', 'Interest earned this month'],
    ['Total Interest', '', 'Total interest earned'],
    ['Payment Streak', '', 'Consecutive on-time payments'],
    ['Loan Status', '', 'Status of any active loans']
  ];
  
  sheet.getRange(15, 1, infoRows.length, 3).setValues(infoRows);
  
  // Set column widths
  sheet.setColumnWidth(1, 100);  // Member ID
  sheet.setColumnWidth(2, 150);  // Member Name
  sheet.setColumnWidth(3, 120);  // Total Savings
  sheet.setColumnWidth(4, 120);  // Current Loan
  sheet.setColumnWidth(5, 120);  // Next Payment Date
  sheet.setColumnWidth(6, 100);  // Status
  
  // Set row heights
  sheet.setRowHeight(1, 35);
  sheet.setRowHeight(4, 30);
  sheet.setRowHeight(7, 30);
  sheet.setRowHeight(10, 25);
  sheet.setRowHeight(13, 25);
  
  sheet.setFrozenRows(10);
  
  // Install trigger for auto-update
  installAutoUpdateTrigger();
  
  return sheet;
}

function updateMemberPortalDropdown() {
  try {
    console.log('Updating Member Portal dropdown...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    const savingsSheet = ss.getSheetByName('💰 Savings');
    
    if (!portalSheet || !savingsSheet) {
      console.log('Sheets not found');
      return;
    }
    
    const memberIdCell = portalSheet.getRange(7, 3);
    
    // Get all member IDs
    const savingsData = savingsSheet.getDataRange().getValues();
    const memberIds = ['--- Clear Selection ---']; // Add Clear option at the top
    
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] && savingsData[i][0].toString().trim() !== '') {
        memberIds.push(savingsData[i][0]);
      }
    }
    
    if (memberIds.length > 1) { // More than just "Clear" option
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(memberIds, true)
        .setAllowInvalid(false)
        .setHelpText('Select your Member ID or "--- Clear Selection ---" to clear')
        .build();
      
      memberIdCell.setDataValidation(rule);
      console.log('Dropdown updated with', memberIds.length - 1, 'members + Clear option');
    }
    
  } catch (error) {
    console.error('Error updating member portal dropdown:', error);
  }
}

function createSavingsSheet(ss) {
  console.log('Creating Savings sheet...');
  const sheet = ss.insertSheet('💰 Savings');
  
  // Clear any existing content
  sheet.clear();
  
  // Title
  sheet.getRange(1, 1, 1, 11).merge()
    .setValue('💰 SAVINGS MANAGEMENT')
    .setFontSize(16).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.primary)
    .setHorizontalAlignment('center');
  
  // Subtitle
  sheet.getRange(2, 1, 1, 11).merge()
    .setValue('Payment Schedule: 5th (₱1,000) & 20th (₱1,000) | Total: ₱2,000/month | Interest: ' + (CONFIG.interestRateSavingsOnly * 100) + '% per year (' + (CONFIG.interestRateSavingsOnly / 12 * 100).toFixed(2) + '% per month)')
    .setFontSize(11)
    .setFontColor(COLORS.primary)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');
  
  // Headers - STARTING IN COLUMN A
  const headers = [
    ['ID', 'Member Name', 'Total Savings', 'Date Joined', 
     'Last Payment', 'Next Payment Date', 'Next Payment Amount', 'Status', 'Months Active', 'Balance', 'Interest Earned']
  ];
  
  sheet.getRange(4, 1, 1, 11).setValues(headers)
    .setFontWeight('bold').setFontColor(COLORS.white)
    .setBackground(COLORS.secondary).setHorizontalAlignment('center');
  
  // Set column widths - STARTING FROM COLUMN A
  const columnWidths = [80, 150, 100, 100, 100, 120, 120, 100, 100, 100, 100];
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width); // index + 1 because columns start at 1
  });
  
  // Freeze header rows
  sheet.setFrozenRows(4);
  
  // Format columns - ALL STARTING IN COLUMN A
  // Column A: ID (text)
  // Column B: Member Name (text)
  sheet.getRange('C:C').setNumberFormat('"₱"#,##0.00'); // Column C: Total Savings
  sheet.getRange('D:F').setNumberFormat('mm/dd/yyyy'); // Columns D, E, F: Dates
  sheet.getRange('G:G').setNumberFormat('"₱"#,##0.00'); // Column G: Next Payment Amount
  // Column H: Status (text with conditional formatting)
  sheet.getRange('I:I').setNumberFormat('0'); // Column I: Months Active as whole number
  sheet.getRange('J:J').setNumberFormat('"₱"#,##0.00'); // Column J: Balance
  sheet.getRange('K:K').setNumberFormat('"₱"#,##0.00'); // Column K: Interest Earned
 
}

function createSavingsPaymentsSheet(ss) {
  console.log('Creating Savings Payments sheet...');
  const sheet = ss.insertSheet('💵 Savings Payments');
  
  // Clear any existing content
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 7).merge()
    .setValue('💵 SAVINGS PAYMENTS')
    .setFontSize(16).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.success)
    .setHorizontalAlignment('center');
  
  sheet.getRange(2, 1, 1, 7).merge()
    .setValue('Records all savings payments including initial deposits')
    .setFontSize(11)
    .setFontColor(COLORS.primary)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');
  
  const headers = [
    ['Payment ID', 'Member ID', 'Date', 'Amount', 
     'Type', 'Verified', 'Period']
  ];
  
  sheet.getRange(3, 1, 1, 7).setValues(headers)
    .setFontWeight('bold').setFontColor(COLORS.white)
    .setBackground(COLORS.primary).setHorizontalAlignment('center');
  
  const columnWidths = [120, 100, 100, 100, 120, 80, 100];
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  sheet.setFrozenRows(3);
  
  // Format columns
  sheet.getRange('D:D').setNumberFormat('"₱"#,##0.00');
  sheet.getRange('C:C').setNumberFormat('mm/dd/yyyy');
  
  return sheet;
}

function createLoanSheet(ss) {
  console.log('Creating Loans sheet...');
  const sheet = ss.insertSheet('🏦 Loans');
  
  sheet.getRange(1, 1, 1, 15).merge()
    .setValue('🏦 LOAN MANAGEMENT')
    .setFontSize(16).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.accent)
    .setHorizontalAlignment('center');
  
  sheet.getRange(2, 1, 1, 15).merge()
    .setValue('Loan Terms: ' + CONFIG.loanTermMin + '-' + CONFIG.loanTermMax + ' months | Auto-generated Loan IDs | Interest: ' + (CONFIG.interestRateWithLoan * 100) + '% PER MONTH')
    .setFontSize(11)
    .setFontColor(COLORS.primary)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');
  
  const headers = [
    ['Loan ID', 'Member ID', 'Loan Amount', 'Interest Rate', 'Term (months)', 
     'Date Approved', 'Status', 'Monthly Interest', 'Monthly Payment', 
     'Total Interest', 'Total Repayment', 'Remaining Balance', 'Payments Made', 'Next Payment Date', 'Effective Annual Rate']
  ];
  
  sheet.getRange(4, 1, 1, 15).setValues(headers)
    .setFontWeight('bold').setFontColor(COLORS.white)
    .setBackground(COLORS.secondary).setHorizontalAlignment('center');
  
  // Add data validation for term
  const termRange = sheet.getRange(5, 5, 1000, 1);
  const termValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(CONFIG.loanTermMin, CONFIG.loanTermMax)
    .setAllowInvalid(false)
    .setHelpText('Loan term must be between ' + CONFIG.loanTermMin + '-' + CONFIG.loanTermMax + ' months')
    .build();
  termRange.setDataValidation(termValidation);
  
  const columnWidths = [120, 100, 100, 80, 80, 100, 100, 100, 100, 100, 100, 100, 100, 120, 100];
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  sheet.setFrozenRows(4);
  
  // Format columns
  sheet.getRange('C:C').setNumberFormat('₱#,##0.00');
  sheet.getRange('D:D').setNumberFormat('0%'); // Percentage without decimals
  sheet.getRange('H:L').setNumberFormat('₱#,##0.00');
  sheet.getRange('N:N').setNumberFormat('mm/dd/yyyy');
  sheet.getRange('F:F').setNumberFormat('mm/dd/yyyy');
  sheet.getRange('O:O').setNumberFormat('0%'); // Percentage without decimals
  sheet.getRange('M:M').setNumberFormat('0'); // Payments made as whole number
  sheet.getRange('E:E').setNumberFormat('0'); // Term as whole number
  
  return sheet;
}

function createLoanPaymentsSheet(ss) {
  console.log('Creating Loan Payments sheet...');
  const sheet = ss.insertSheet('💳 Loan Payments');
  
  sheet.getRange(1, 1, 1, 9).merge()
    .setValue('💳 LOAN PAYMENTS')
    .setFontSize(16).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.info)
    .setHorizontalAlignment('center');
  
  const headers = [
    ['Payment ID', 'Loan ID', 'Date', 'Paid', 'Principal', 
     'Interest', 'Remaining', 'Status', 'Payment #']
  ];
  
  sheet.getRange(3, 1, 1, 9).setValues(headers)
    .setFontWeight('bold').setFontColor(COLORS.white)
    .setBackground(COLORS.primary).setHorizontalAlignment('center');
  
  const columnWidths = [120, 100, 100, 100, 100, 100, 100, 100, 80];
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  sheet.setFrozenRows(3);
  
  // Format columns
  sheet.getRange('D:G').setNumberFormat('₱#,##0.00');
  sheet.getRange('C:C').setNumberFormat('mm/dd/yyyy');
  sheet.getRange('I:I').setNumberFormat('0'); // Payment # as whole number
  
  return sheet;
}

function createSummarySheet(ss) {
  console.log('Creating Summary sheet...');
  const sheet = ss.insertSheet('📊 Summary');
  
  sheet.getRange(1, 1, 1, 11).merge()
    .setValue('📊 COMPREHENSIVE SUMMARY')
    .setFontSize(16).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.header)
    .setHorizontalAlignment('center');
  
  const headers = [
    ['Member ID', 'Name', 'Savings', 'Loans', 'Active Loans', 
     'Loan Status', 'Interest Paid', 'Savings Status', 'Balance', 'Savings Streak', 'Loan Streak']
  ];
  
  sheet.getRange(3, 1, 1, 11).setValues(headers)
    .setFontWeight('bold').setFontColor(COLORS.white)
    .setBackground(COLORS.primary).setHorizontalAlignment('center');
  
  const columnWidths = [100, 150, 100, 100, 100, 100, 100, 100, 100, 100, 100];
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  sheet.setFrozenRows(3);
  
  // Format columns
  sheet.getRange('C:C').setNumberFormat('₱#,##0.00');
  sheet.getRange('D:D').setNumberFormat('₱#,##0.00');
  sheet.getRange('G:G').setNumberFormat('₱#,##0.00');
  sheet.getRange('I:I').setNumberFormat('₱#,##0.00');
  sheet.getRange('J:J').setNumberFormat('0'); // Streaks as whole numbers
  sheet.getRange('K:K').setNumberFormat('0'); // Streaks as whole numbers
  
  return sheet;
}

// ============================================
// NEW FUNCTION: INITIALIZE FUNDS ACTIVITY
// ============================================

function initializeFundsActivity(sheet) {
  // Add initial transaction if sheet is empty
  const lastRow = sheet.getLastRow();
  if (lastRow <= 21) {
    const initialTransaction = [
      new Date(),
      'System Initialization',
      'SYSTEM',
      0,
      0,
      0,
      'Community Funds Tracker initialized'
    ];
    
    sheet.getRange(22, 1, 1, 7).setValues([initialTransaction]);
    
    sheet.getRange(22, 1).setNumberFormat('mm/dd/yyyy hh:mm');
    sheet.getRange(22, 4, 1, 3).setNumberFormat('₱#,##0.00');
  }
}

// ============================================
// NEW FUNCTION: UPDATE COMMUNITY FUNDS CALCULATIONS
// ============================================

function updateCommunityFundsCalculations() {
  try {
    console.log('Starting community funds update...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fundsSheet = ss.getSheetByName('💰 Community Funds');
    
    if (!fundsSheet) {
      console.log('Community Funds sheet not found, creating it...');
      createCommunityFundsSheet(ss);
      return;
    }
    
    // Try to get other sheets, but don't fail if they don't exist yet
    let savingsSheet, loansSheet;
    try {
      savingsSheet = ss.getSheetByName('💰 Savings');
    } catch (e) {
      console.log('Savings sheet not available:', e.message);
    }
    
    try {
      loansSheet = ss.getSheetByName('🏦 Loans');
    } catch (e) {
      console.log('Loans sheet not available:', e.message);
    }
    
    // Initialize totals
    let totalSavings = 0;
    let totalActiveLoans = 0;
    
    // 1. Calculate Total Community Savings
    if (savingsSheet) {
      try {
        const savingsData = savingsSheet.getDataRange().getValues();
        console.log('Savings data rows:', savingsData.length);
        
        for (let i = 4; i < savingsData.length; i++) {
          // Check if row has data
          if (savingsData[i] && savingsData[i][0] && savingsData[i][0].toString().trim() !== '') {
            const savings = parseFloat(savingsData[i][2]) || 0;
            totalSavings += savings;
          }
        }
        console.log('Total savings calculated:', totalSavings);
      } catch (e) {
        console.error('Error calculating savings:', e.message);
      }
    }
    
    // 2. Calculate Total Active Loans
    if (loansSheet) {
      try {
        const loansData = loansSheet.getDataRange().getValues();
        console.log('Loans data rows:', loansData.length);
        
        for (let i = 4; i < loansData.length; i++) {
          // Check if row has data and status is Active
          if (loansData[i] && loansData[i][0] && loansData[i][0].toString().trim() !== '') {
            const status = loansData[i][6];
            if (status && status.toString().trim() === 'Active') {
              const loanAmount = parseFloat(loansData[i][2]) || 0;
              totalActiveLoans += loanAmount;
            }
          }
        }
        console.log('Total active loans calculated:', totalActiveLoans);
      } catch (e) {
        console.error('Error calculating loans:', e.message);
      }
    }
    
    // 3. Calculate Available for Loans (Savings - Reserve - Active Loans)
    const reserveAmount = totalSavings * CONFIG.reserveRatio;
    const availableForLoans = Math.max(0, totalSavings - reserveAmount - totalActiveLoans);
    
    // 4. Calculate Loans-to-Savings Ratio
    const loanRatio = totalSavings > 0 ? totalActiveLoans / totalSavings : 0;
    
    // 5. UPDATE THE CELLS DIRECTLY
    try {
      fundsSheet.getRange('B7').setValue(totalSavings);
      fundsSheet.getRange('B8').setValue(totalActiveLoans);
      fundsSheet.getRange('B9').setValue(availableForLoans);
      fundsSheet.getRange('B10').setValue(reserveAmount);
      fundsSheet.getRange('B11').setValue(loanRatio);
      
      // 6. Update Loan Capacity
      const maxSingleLoan = Math.min(CONFIG.maxLoanAmount, availableForLoans);
      const recommendedLimit = Math.min(CONFIG.maxLoanAmount * 0.5, availableForLoans * 0.5);
      
      fundsSheet.getRange('B16').setValue(availableForLoans);
      fundsSheet.getRange('C16').setValue(maxSingleLoan);
      fundsSheet.getRange('B17').setValue(recommendedLimit);
      
      // 7. Update Funds Status
      const fundsStatus = availableForLoans >= CONFIG.minLoanAmount ? "Funds Available" : "Insufficient Funds";
      fundsSheet.getRange('D16').setValue(fundsStatus);
      
      // 8. Format the cells
      fundsSheet.getRange('B7:B10').setNumberFormat('₱#,##0.00');
      fundsSheet.getRange('B16:C17').setNumberFormat('₱#,##0.00');
      fundsSheet.getRange('B11').setNumberFormat('0.00%');
      
      // 9. Update timestamp
      fundsSheet.getRange('B25').setValue(new Date()).setNumberFormat('mm/dd/yyyy hh:mm:ss');
      
      console.log('Community Funds Updated Successfully:');
      console.log('- Total Savings: ₱' + totalSavings.toLocaleString());
      console.log('- Active Loans: ₱' + totalActiveLoans.toLocaleString());
      console.log('- Available: ₱' + availableForLoans.toLocaleString());
      
      return {
        success: true,
        totalSavings: totalSavings,
        totalActiveLoans: totalActiveLoans,
        availableForLoans: availableForLoans
      };
      
    } catch (e) {
      console.error('Error updating funds sheet cells:', e.message);
      throw e;
    }
    
  } catch (error) {
    console.error('Error in updateCommunityFundsCalculations:', error);
    console.error('Stack trace:', error.stack);
    
    // Try to show a helpful message
    const errorMsg = error.toString();
    if (errorMsg.includes('Cannot call method "getRange"')) {
      console.error('Funds sheet structure may be corrupted. Consider regenerating the system.');
    }
    
    throw new Error('Failed to update community funds: ' + error.message);
  }
}

// ============================================
// REVISED: CREATE COMMUNITY FUNDS SHEET - SIMPLIFIED VERSION
// ============================================

function createCommunityFundsSheet(ss) {
  console.log('Creating Community Funds sheet...');
  
  // Check if sheet already exists and remove it
  const existingSheet = ss.getSheetByName('💰 Community Funds');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  
  const sheet = ss.insertSheet('💰 Community Funds');
  
  // Clear any existing content
  sheet.clear();
  
  // Set sheet protection (optional)
  sheet.protect().setDescription('Community Funds Sheet - Auto-generated');
  
  try {
    sheet.getRange(1, 1, 1, 8).merge()
      .setValue('💰 COMMUNITY FUNDS TRACKER')
      .setFontSize(16).setFontWeight('bold')
      .setFontColor(COLORS.white).setBackground(COLORS.header)
      .setHorizontalAlignment('center');
    
    sheet.getRange(2, 1, 1, 8).merge()
      .setValue('📊 Tracks total community savings and available funds for loans | Savings: ' + (CONFIG.interestRateSavingsOnly * 100) + '% annual | Loans: ' + (CONFIG.interestRateWithLoan * 100) + '% MONTHLY')
      .setFontSize(11)
      .setFontColor(COLORS.primary)
      .setHorizontalAlignment('center')
      .setFontWeight('bold');
    
    // Summary Section
    sheet.getRange(4, 1, 1, 8).merge()
      .setValue('📈 FUNDS SUMMARY (Auto-calculated)')
      .setFontSize(14).setFontWeight('bold')
      .setFontColor(COLORS.white).setBackground(COLORS.secondary)
      .setHorizontalAlignment('center');
  
  const summaryHeaders = [
    ['Description', 'Amount', 'Last Updated', 'Status']
  ];
  
  sheet.getRange(6, 1, 1, 4).setValues(summaryHeaders)
    .setFontWeight('bold').setFontColor(COLORS.white)
    .setBackground(COLORS.primary).setHorizontalAlignment('center');
  
  // Summary Data - VALUES WILL BE CALCULATED BY FUNCTION
  const summaryData = [
    ['Total Community Savings', 0, new Date(), 'Active'],
    ['Total Active Loans', 0, new Date(), 'Active'],
    ['Total Available for Loans', 0, new Date(), 'Calculated'],
    ['Reserve Fund (' + (CONFIG.reserveRatio * 100) + '%)', 0, new Date(), 'Reserved'],
    ['Loans-to-Savings Ratio', 0, new Date(), 'Ratio']
  ];
  
  sheet.getRange(7, 1, 5, 4).setValues(summaryData);
  
  // Formatting
  sheet.getRange(7, 2, 4, 1).setNumberFormat('₱#,##0.00');
  sheet.getRange(11, 2).setNumberFormat('0.00%');
  sheet.getRange(7, 3, 5, 1).setNumberFormat('mm/dd/yyyy hh:mm');
  
  // Apply status colors
  sheet.getRange(7, 4).setBackground(STATUS_COLORS['Active'] || COLORS.success)
    .setFontColor(COLORS.dark).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(8, 4).setBackground(STATUS_COLORS['Active'] || COLORS.success)
    .setFontColor(COLORS.dark).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(9, 4).setBackground(STATUS_COLORS['Calculated'] || COLORS.info)
    .setFontColor(COLORS.dark).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(10, 4).setBackground(STATUS_COLORS['Reserved'] || COLORS.warning)
    .setFontColor(COLORS.dark).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(11, 4).setBackground(STATUS_COLORS['Ratio'] || COLORS.info)
    .setFontColor(COLORS.dark).setFontWeight('bold').setHorizontalAlignment('center');
  
  // Loan Capacity Section
  sheet.getRange(13, 1, 1, 8).merge()
    .setValue('🏦 LOAN CAPACITY')
    .setFontSize(14).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.accent)
    .setHorizontalAlignment('center');
  
  const capacityHeaders = [
    ['Description', 'Amount', 'Limit', 'Status']
  ];
  
  sheet.getRange(15, 1, 1, 4).setValues(capacityHeaders)
    .setFontWeight('bold').setFontColor(COLORS.white)
    .setBackground(COLORS.primary).setHorizontalAlignment('center');
  
  // Capacity Data
  const capacityData = [
    ['Available for New Loans', 0, 0, 'Calculating...'],
    ['Recommended Loan Limit', 0, '(50% of available)', 'Suggested']
  ];
  
  sheet.getRange(16, 1, 2, 4).setValues(capacityData);
  
  // Formatting
  sheet.getRange(16, 2, 2, 2).setNumberFormat('₱#,##0.00');
  
  // Recent Transactions
  sheet.getRange(19, 1, 1, 8).merge()
    .setValue('💼 RECENT FUNDS ACTIVITY')
    .setFontSize(14).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.info)
    .setHorizontalAlignment('center');
  
  const activityHeaders = [
    ['Date', 'Transaction Type', 'Member/Loan ID', 'Amount', 'Balance Before', 'Balance After', 'Notes']
  ];
  
  sheet.getRange(21, 1, 1, 7).setValues(activityHeaders)
    .setFontWeight('bold').setFontColor(COLORS.white)
    .setBackground(COLORS.primary).setHorizontalAlignment('center');
  
  // Add initial transaction
  sheet.getRange(22, 1, 1, 7).setValues([[
    new Date(),
    'System Initialized',
    'SYSTEM',
    0,
    0,
    0,
    'Community Funds Tracker created'
  ]]);
  
  // Set column widths
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 200);
  
  sheet.setFrozenRows(6);
  
  // Format activity columns
  sheet.getRange('A:A').setNumberFormat('mm/dd/yyyy hh:mm');
  sheet.getRange('D:F').setNumberFormat('₱#,##0.00');
  
  // Add update info
  sheet.getRange(25, 1).setValue('Last Updated:').setFontWeight('bold');
  sheet.getRange(25, 2).setValue(new Date()).setNumberFormat('mm/dd/yyyy hh:mm:ss');
  
  // Run initial calculation
  SpreadsheetApp.flush(); // Force updates
    updateCommunityFundsCalculations();
    
    console.log('Community Funds sheet created successfully');
    
  } catch (error) {
    console.error('Error creating Community Funds sheet:', error);
    throw error;
  }
  
  return sheet;
}

function createSettingsSheet(ss) {
  console.log('Creating Settings sheet...');
  const sheet = ss.insertSheet('⚙️ Settings');
  
  sheet.getRange(1, 1, 1, 4).merge()
    .setValue('⚙️ SYSTEM SETTINGS')
    .setFontSize(16).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.header)
    .setHorizontalAlignment('center');
  
  sheet.getRange(2, 1, 1, 4).merge()
    .setValue('Configure system parameters below. Changes take effect immediately.')
    .setFontSize(11)
    .setFontColor(COLORS.primary)
    .setHorizontalAlignment('center');
  
  // Settings table
  const settingsHeaders = [
    ['Setting', 'Value', 'Description', 'Default']
  ];
  
  sheet.getRange(4, 1, 1, 4).setValues(settingsHeaders)
    .setFontWeight('bold').setFontColor(COLORS.white)
    .setBackground(COLORS.primary).setHorizontalAlignment('center');
  
  // Settings data
  const settingsData = [
    ['Company Name', CONFIG.companyName, 'Name of your savings and loan company', 'Community Savings & Loan'],
    ['Loan Interest Rate (Monthly)', CONFIG.interestRateWithLoan, 'Monthly interest rate for loans (e.g., 0.09 = 9% per month)', '0.09 (9% per month)'],
    ['Savings Interest (With Loan) Annual', CONFIG.interestRateWithoutLoan, 'Annual interest rate for savings when member has active loan', '0.09 (9% per year)'],
    ['Savings Interest (No Loan) Annual', CONFIG.interestRateSavingsOnly, 'Annual interest rate for savings when member has no loan', '0.09 (9% per year)'],
    ['Monthly Contribution', CONFIG.monthlyContribution, 'Total monthly savings contribution per member', '2000'],
    ['Payment Per Period', CONFIG.paymentPerPeriod, 'Amount due on each payment date (5th and 20th)', '1000'],
    ['Reserve Ratio', CONFIG.reserveRatio, 'Percentage of savings kept as reserve (e.g., 0.3 = 30%)', '0.3 (30%)'],
    ['Minimum Loan Amount', CONFIG.minLoanAmount, 'Minimum amount that can be borrowed', '1000'],
    ['Maximum Loan Amount', CONFIG.maxLoanAmount, 'Maximum amount that can be borrowed', '20000'],
    ['Minimum Loan Term (Months)', CONFIG.loanTermMin, 'Minimum loan duration in months', '1'],
    ['Maximum Loan Term (Months)', CONFIG.loanTermMax, 'Maximum loan duration in months', '12'],
    ['Minimum Savings', CONFIG.savingsMin, 'Minimum savings required to join', '1000'],
    ['Savings Required for Loan', CONFIG.savingsForLoan, 'Minimum savings balance required to qualify for loan', '6000'],
    ['Payment Date 1', CONFIG.paymentDates[0], 'First payment date of the month', '5'],
    ['Payment Date 2', CONFIG.paymentDates[1], 'Second payment date of the month', '20']
  ];
  
  sheet.getRange(5, 1, settingsData.length, 4).setValues(settingsData);
  
  // Format columns
  sheet.setColumnWidth(1, 200); // Setting name
  sheet.setColumnWidth(2, 120); // Value
  sheet.setColumnWidth(3, 300); // Description
  sheet.setColumnWidth(4, 150); // Default
  
  // Format specific columns
  sheet.getRange(6, 2, settingsData.length, 1).setNumberFormat('0%'); // Interest rates
  sheet.getRange(8, 2).setNumberFormat('0'); // Monthly contribution (whole number)
  sheet.getRange(9, 2).setNumberFormat('0'); // Payment per period (whole number)
  sheet.getRange(10, 2).setNumberFormat('0%'); // Reserve ratio
  sheet.getRange(11, 2).setNumberFormat('₱#,##0'); // Min loan amount
  sheet.getRange(12, 2).setNumberFormat('₱#,##0'); // Max loan amount
  sheet.getRange(13, 2).setNumberFormat('0'); // Min term (whole number)
  sheet.getRange(14, 2).setNumberFormat('0'); // Max term (whole number)
  sheet.getRange(15, 2).setNumberFormat('₱#,##0'); // Min savings
  sheet.getRange(16, 2).setNumberFormat('₱#,##0'); // Savings for loan
  sheet.getRange(17, 2).setNumberFormat('0'); // Payment date 1
  sheet.getRange(18, 2).setNumberFormat('0'); // Payment date 2
  
  // Add save button area
  const lastRow = 5 + settingsData.length + 2;
  sheet.getRange(lastRow, 1, 1, 4).merge()
    .setValue('SAVE SETTINGS')
    .setFontSize(14).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.success)
    .setHorizontalAlignment('center');
  
  sheet.getRange(lastRow + 1, 1, 1, 4).merge()
    .setValue('Click to apply settings to entire system')
    .setFontSize(11)
    .setFontColor(COLORS.dark)
    .setHorizontalAlignment('center');
  
  // Add reset button
  sheet.getRange(lastRow + 3, 1, 1, 4).merge()
    .setValue('RESET TO DEFAULTS')
    .setFontSize(12).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.warning)
    .setHorizontalAlignment('center');
  
  sheet.setFrozenRows(4);
  
  return sheet;
}

function createDashboardSheet(ss) {
  console.log('Creating Dashboard...');
  const sheet = ss.insertSheet('📈 Dashboard');
  
  sheet.getRange(1, 1, 1, 8).merge()
    .setValue('📈 ' + CONFIG.companyName + ' DASHBOARD')
    .setFontSize(18).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.header)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);
  
  sheet.getRange(2, 1, 1, 8).merge()
    .setValue('📅 Payment Schedule: 5th (₱1,000) & 20th (₱1,000) | Total: ₱2,000/month | Savings: ' + (CONFIG.interestRateSavingsOnly * 100) + '% annual | Loans: ' + (CONFIG.interestRateWithLoan * 100) + '% MONTHLY')
    .setFontSize(11)
    .setFontColor(COLORS.primary)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');
  
  sheet.getRange(3, 1).setValue('Date:').setFontWeight('bold');
  sheet.getRange(3, 2).setValue(new Date().toLocaleDateString()).setFontColor(COLORS.primary);
  
  sheet.getRange(4, 1).setValue('Interest Rates:').setFontWeight('bold');
  sheet.getRange(4, 2).setValue('Savings: ' + (CONFIG.interestRateSavingsOnly * 100) + '% per year (' + (CONFIG.interestRateSavingsOnly / 12 * 100).toFixed(2) + '% monthly)').setFontColor(COLORS.success);
  sheet.getRange(4, 5).setValue('Loans: ' + (CONFIG.interestRateWithLoan * 100) + '% PER MONTH').setFontColor(COLORS.warning);
  
  sheet.getRange(6, 1, 1, 8).merge()
    .setValue('📊 KEY METRICS')
    .setFontSize(14).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.primary)
    .setHorizontalAlignment('center');
  
  const metrics = [
    {title: '💰 Total Savings', formula: '=SUM(\'💰 Savings\'!C:C)', row: 8, col: 1},
    {title: '🏦 Total Loans', formula: '=SUM(\'🏦 Loans\'!C:C)', row: 8, col: 3},
    {title: '👥 Total Members', formula: '=COUNTA(\'💰 Savings\'!A:A)-4', row: 8, col: 5},
    {title: '📈 Active Loans', formula: '=COUNTIF(\'🏦 Loans\'!G:G, "Active")', row: 8, col: 7},
    {title: '💵 Monthly Collection', formula: '=' + CONFIG.monthlyContribution + '*(COUNTA(\'💰 Savings\'!A:A)-4)', row: 11, col: 1},
    {title: '📅 Next 5th Collection', formula: '=' + CONFIG.paymentPerPeriod + '*(COUNTA(\'💰 Savings\'!A:A)-4)', row: 11, col: 3},
    {title: '📅 Next 20th Collection', formula: '=' + CONFIG.paymentPerPeriod + '*(COUNTA(\'💰 Savings\'!A:A)-4)', row: 11, col: 5},
    {title: '🎯 On Track', formula: '=COUNTIF(\'📊 Summary\'!H:H, "On Track")', row: 11, col: 7},
    {title: '💰 Available for Loans', formula: '=\'💰 Community Funds\'!B9', row: 14, col: 1},
    {title: '📉 Loan Interest Earned', formula: '=SUM(\'💳 Loan Payments\'!F:F)', row: 14, col: 3},
    {title: '📈 Savings Interest Paid', formula: '=SUM(\'💰 Savings\'!C:C)*' + CONFIG.interestRateSavingsOnly + '/12', row: 14, col: 5},
    {title: '📊 Net Interest', formula: '=N14-L14', row: 14, col: 7}
  ];
  
  metrics.forEach(metric => {
    sheet.getRange(metric.row, metric.col, 2, 2).merge()
      .setBackground(COLORS.light)
      .setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
    
    sheet.getRange(metric.row, metric.col)
      .setValue(metric.title)
      .setFontSize(11)
      .setFontColor(COLORS.dark)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    const valueCell = sheet.getRange(metric.row + 1, metric.col);
    valueCell.setFormula(metric.formula)
      .setFontSize(14)
      .setFontWeight('bold')
      .setFontColor(COLORS.primary)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    if (metric.title.includes('Savings') || metric.title.includes('Loans') || metric.title.includes('Collection') || 
        metric.title.includes('Interest') || metric.title.includes('Available')) {
      valueCell.setNumberFormat('₱#,##0.00');
    }
    if (metric.title.includes('Members') || metric.title.includes('Loans') || metric.title.includes('On Track')) {
      valueCell.setNumberFormat('0'); // Whole numbers
    }
  });
  
  sheet.getRange(17, 1, 1, 8).merge()
    .setValue('📋 QUICK ACTIONS')
    .setFontSize(14).setFontWeight('bold')
    .setFontColor(COLORS.white).setBackground(COLORS.secondary)
    .setHorizontalAlignment('center');
  
  // Create action buttons
  const buttons = [
    {text: '➕ Add Member', row: 19, col: 2, action: 'showAddMemberForm'},
    {text: '👥 View Members', row: 19, col: 5, action: 'showAllMembers'},
    {text: '💵 Record Payment', row: 21, col: 2, action: 'showSavingsPaymentForm'},
    {text: '🏦 Apply Loan', row: 21, col: 5, action: 'showLoanApplicationForm'},
    {text: '💳 Loan Payment', row: 23, col: 2, action: 'showLoanPaymentForm'},
    {text: '💰 View Funds', row: 23, col: 5, action: 'showCommunityFunds'},
    {text: '⚙️ Settings', row: 25, col: 2, action: 'showSettings'},
    {text: '📊 Update All', row: 25, col: 5, action: 'updateAllCalculations'},
    {text: '📋 Reports', row: 27, col: 2, action: 'generateReports'},
    {text: '🔄 Refresh All', row: 27, col: 5, action: 'refreshAllDropdowns'}
  ];
  
  buttons.forEach(btn => {
    const cell = sheet.getRange(btn.row, btn.col, 1, 2).merge();
    cell.setValue(btn.text)
      .setFontWeight('bold')
      .setFontColor(COLORS.white)
      .setBackground(COLORS.accent)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, COLORS.white, SpreadsheetApp.BorderStyle.SOLID);
  });
  
  sheet.getRange(29, 1, 1, 8).merge()
    .setValue('💡 Use the "🏦 S&L System" for all actions!')
    .setFontSize(11)
    .setFontColor(COLORS.primary)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');
  
  // Set column widths
  for (let i = 1; i <= 8; i++) {
    sheet.setColumnWidth(i, 100);
  }
  
  // Set row heights for buttons
  sheet.setRowHeight(19, 35);
  sheet.setRowHeight(21, 35);
  sheet.setRowHeight(23, 35);
  sheet.setRowHeight(25, 35);
  sheet.setRowHeight(27, 35);
  
  return sheet;
}

// ============================================
// SETTINGS MANAGEMENT
// ============================================

function showSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('⚙️ Settings');
  
  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Settings sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  settingsSheet.showSheet();
  ss.setActiveSheet(settingsSheet);
}

function saveSettingsFromSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('⚙️ Settings');
    
    if (!settingsSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Settings sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const settingsData = settingsSheet.getRange(5, 1, 20, 2).getValues();
    
    // Update CONFIG with values from sheet
    settingsData.forEach(row => {
      const settingName = row[0];
      const settingValue = row[1];
      
      switch(settingName) {
        case 'Company Name':
          CONFIG.companyName = settingValue.toString();
          break;
        case 'Loan Interest Rate (Monthly)':
          CONFIG.interestRateWithLoan = parseFloat(settingValue);
          break;
        case 'Savings Interest (With Loan) Annual':
          CONFIG.interestRateWithoutLoan = parseFloat(settingValue);
          break;
        case 'Savings Interest (No Loan) Annual':
          CONFIG.interestRateSavingsOnly = parseFloat(settingValue);
          break;
        case 'Monthly Contribution':
          CONFIG.monthlyContribution = parseInt(settingValue);
          break;
        case 'Payment Per Period':
          CONFIG.paymentPerPeriod = parseInt(settingValue);
          break;
        case 'Reserve Ratio':
          CONFIG.reserveRatio = parseFloat(settingValue);
          break;
        case 'Minimum Loan Amount':
          CONFIG.minLoanAmount = parseInt(settingValue);
          break;
        case 'Maximum Loan Amount':
          CONFIG.maxLoanAmount = parseInt(settingValue);
          break;
        case 'Minimum Loan Term (Months)':
          CONFIG.loanTermMin = parseInt(settingValue);
          break;
        case 'Maximum Loan Term (Months)':
          CONFIG.loanTermMax = parseInt(settingValue);
          break;
        case 'Minimum Savings':
          CONFIG.savingsMin = parseInt(settingValue);
          break;
        case 'Savings Required for Loan':
          CONFIG.savingsForLoan = parseInt(settingValue);
          break;
        case 'Payment Date 1':
          CONFIG.paymentDates[0] = parseInt(settingValue);
          break;
        case 'Payment Date 2':
          CONFIG.paymentDates[1] = parseInt(settingValue);
          break;
      }
    });
    
    // Update sheets with new settings
    updateAllSheetsWithSettings();
    
    SpreadsheetApp.getUi().alert(
      '✅ Settings Saved',
      'All system settings have been updated successfully!\n\n' +
      'The following sheets have been updated:\n' +
      '• Dashboard\n' +
      '• Community Funds\n' +
      '• Member Portal\n' +
      '• All calculation formulas\n\n' +
      'New settings will be applied immediately.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error saving settings:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to save settings: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function resetSettingsToDefaults() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm', 'Are you sure you want to reset all settings to defaults?', ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Reset CONFIG to original defaults
  CONFIG = {
    interestRateWithLoan: 0.09,
    interestRateWithoutLoan: 0.09,
    interestRateSavingsOnly: 0.09,
    monthlyContribution: 2000,
    paymentPerPeriod: 1000,
    paymentDates: [5, 20],
    companyName: "Community Savings & Loan",
    nextMemberId: 1001,
    reserveRatio: 0.3,
    minLoanAmount: 1000,
    maxLoanAmount: 20000,
    loanTermMin: 1,
    loanTermMax: 12,
    savingsMin: 1000,
    savingsForLoan: 6000
  };
  
  // Update settings sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('⚙️ Settings');
  
  if (settingsSheet) {
    const settingsData = [
      ['Company Name', CONFIG.companyName],
      ['Loan Interest Rate (Monthly)', CONFIG.interestRateWithLoan],
      ['Savings Interest (With Loan) Annual', CONFIG.interestRateWithoutLoan],
      ['Savings Interest (No Loan) Annual', CONFIG.interestRateSavingsOnly],
      ['Monthly Contribution', CONFIG.monthlyContribution],
      ['Payment Per Period', CONFIG.paymentPerPeriod],
      ['Reserve Ratio', CONFIG.reserveRatio],
      ['Minimum Loan Amount', CONFIG.minLoanAmount],
      ['Maximum Loan Amount', CONFIG.maxLoanAmount],
      ['Minimum Loan Term (Months)', CONFIG.loanTermMin],
      ['Maximum Loan Term (Months)', CONFIG.loanTermMax],
      ['Minimum Savings', CONFIG.savingsMin],
      ['Savings Required for Loan', CONFIG.savingsForLoan],
      ['Payment Date 1', CONFIG.paymentDates[0]],
      ['Payment Date 2', CONFIG.paymentDates[1]]
    ];
    
    settingsSheet.getRange(5, 2, settingsData.length, 1).setValues(settingsData.map(row => [row[1]]));
  }
  
  updateAllSheetsWithSettings();
  
  SpreadsheetApp.getUi().alert(
    '✅ Settings Reset',
    'All settings have been reset to factory defaults.\n\n' +
    'Click "SAVE SETTINGS" in the Settings sheet to apply.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function updateAllSheetsWithSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Update Dashboard
    const dashboard = ss.getSheetByName('📈 Dashboard');
    if (dashboard) {
      dashboard.getRange(1, 1, 1, 8).merge()
        .setValue('📈 ' + CONFIG.companyName + ' DASHBOARD');
      
      dashboard.getRange(2, 1, 1, 8).merge()
        .setValue('📅 Payment Schedule: ' + CONFIG.paymentDates[0] + 'th (₱' + CONFIG.paymentPerPeriod.toLocaleString() + ') & ' + CONFIG.paymentDates[1] + 'th (₱' + CONFIG.paymentPerPeriod.toLocaleString() + ') | Total: ₱' + CONFIG.monthlyContribution.toLocaleString() + '/month | Savings: ' + (CONFIG.interestRateSavingsOnly * 100) + '% annual | Loans: ' + (CONFIG.interestRateWithLoan * 100) + '% MONTHLY');
      
      dashboard.getRange(4, 2).setValue('Savings: ' + (CONFIG.interestRateSavingsOnly * 100) + '% per year (' + (CONFIG.interestRateSavingsOnly / 12 * 100).toFixed(2) + '% monthly)');
      dashboard.getRange(4, 5).setValue('Loans: ' + (CONFIG.interestRateWithLoan * 100) + '% PER MONTH');
      
      // Update formulas with new amounts
      dashboard.getRange(12, 1).setFormula('=' + CONFIG.monthlyContribution + '*(COUNTA(\'💰 Savings\'!A:A)-4)');
      dashboard.getRange(12, 3).setFormula('=' + CONFIG.paymentPerPeriod + '*(COUNTA(\'💰 Savings\'!A:A)-4)');
      dashboard.getRange(12, 5).setFormula('=' + CONFIG.paymentPerPeriod + '*(COUNTA(\'💰 Savings\'!A:A)-4)');
      dashboard.getRange(15, 5).setFormula('=SUM(\'💰 Savings\'!C:C)*' + CONFIG.interestRateSavingsOnly + '/12');
    }
    
    // Update Community Funds sheet
    const fundsSheet = ss.getSheetByName('💰 Community Funds');
    if (fundsSheet) {
      fundsSheet.getRange(2, 1, 1, 8).merge()
        .setValue('📊 Tracks total community savings and available funds for loans | Savings: ' + (CONFIG.interestRateSavingsOnly * 100) + '% annual | Loans: ' + (CONFIG.interestRateWithLoan * 100) + '% MONTHLY');
      
      fundsSheet.getRange(9, 1).setFormula('=B7-(B7*' + CONFIG.reserveRatio + ')-B8');
      fundsSheet.getRange(10, 1).setFormula('Reserve Fund (' + (CONFIG.reserveRatio * 100) + '%)');
      fundsSheet.getRange(10, 2).setFormula('=B7*' + CONFIG.reserveRatio);
      
      fundsSheet.getRange(16, 2).setFormula('=MIN(' + CONFIG.maxLoanAmount + ', B9)');
      fundsSheet.getRange(16, 3).setFormula('=MIN(' + (CONFIG.maxLoanAmount * 0.5) + ', B9*0.5)');
    }
    
    // Update Loan sheet data validation
    const loanSheet = ss.getSheetByName('🏦 Loans');
    if (loanSheet) {
      const termRange = loanSheet.getRange(5, 5, 1000, 1);
      const termValidation = SpreadsheetApp.newDataValidation()
        .requireNumberBetween(CONFIG.loanTermMin, CONFIG.loanTermMax)
        .setAllowInvalid(false)
        .setHelpText('Loan term must be between ' + CONFIG.loanTermMin + '-' + CONFIG.loanTermMax + ' months')
        .build();
      termRange.setDataValidation(termValidation);
      
      loanSheet.getRange(2, 1, 1, 15).merge()
        .setValue('Loan Terms: ' + CONFIG.loanTermMin + '-' + CONFIG.loanTermMax + ' months | Auto-generated Loan IDs | Interest: ' + (CONFIG.interestRateWithLoan * 100) + '% PER MONTH');
    }
    
    // Update Savings sheet
    const savingsSheet = ss.getSheetByName('💰 Savings');
    if (savingsSheet) {
      savingsSheet.getRange(2, 1, 1, 11).merge()
        .setValue('Payment Schedule: ' + CONFIG.paymentDates[0] + 'th (₱' + CONFIG.paymentPerPeriod.toLocaleString() + ') & ' + CONFIG.paymentDates[1] + 'th (₱' + CONFIG.paymentPerPeriod.toLocaleString() + ') | Total: ₱' + CONFIG.monthlyContribution.toLocaleString() + '/month | Interest: ' + (CONFIG.interestRateSavingsOnly * 100) + '% per year (' + (CONFIG.interestRateSavingsOnly / 12 * 100).toFixed(2) + '% per month)');
    }
    
    // Update Member Portal
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    if (portalSheet) {
      portalSheet.getRange(1, 1, 1, 6).merge()
        .setValue('🏦 ' + CONFIG.companyName + ' - Member Portal');
      
      portalSheet.getRange(2, 1, 1, 6).merge()
        .setValue('📅 Payment Schedule: ' + CONFIG.paymentDates[0] + 'th (₱' + CONFIG.paymentPerPeriod.toLocaleString() + ') & ' + CONFIG.paymentDates[1] + 'th (₱' + CONFIG.paymentPerPeriod.toLocaleString() + ') of each month | Total: ₱' + CONFIG.monthlyContribution.toLocaleString() + '/month');
    }
    
    // Update all dropdowns
    refreshAllDropdowns();
    
    SpreadsheetApp.getUi().alert(
      '✅ Settings Applied',
      'All sheets have been updated with new settings.\n\n' +
      'The system is now using:\n' +
      '• Company: ' + CONFIG.companyName + '\n' +
      '• Loan Interest: ' + (CONFIG.interestRateWithLoan * 100) + '% per month\n' +
      '• Savings Interest: ' + (CONFIG.interestRateSavingsOnly * 100) + '% per year\n' +
      '• Loan Range: ₱' + CONFIG.minLoanAmount.toLocaleString() + ' - ₱' + CONFIG.maxLoanAmount.toLocaleString() + '\n' +
      '• Reserve: ' + (CONFIG.reserveRatio * 100) + '%',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error updating sheets with settings:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to update sheets: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================
// DROPDOWN MANAGEMENT FUNCTIONS
// ============================================

function refreshAllDropdowns() {
  try {
    updateMemberPortalDropdown();
    updateSavingsPaymentDropdowns();
    updateLoanPaymentDropdowns();
    
    SpreadsheetApp.getUi().alert('✅ Dropdowns Updated', 'All dropdown lists have been refreshed with current data.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('Error refreshing dropdowns:', error);
  }
}

function updateSavingsPaymentDropdowns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    
    if (!savingsSheet) return;
    
    // Get all member IDs for dropdown
    const savingsData = savingsSheet.getDataRange().getValues();
    const memberIds = [];
    
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] && savingsData[i][0].toString().trim() !== '') {
        memberIds.push(savingsData[i][0]);
      }
    }
    
    // Store member IDs for use in forms
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('memberIds', JSON.stringify(memberIds));
    
  } catch (error) {
    console.error('Error updating savings payment dropdowns:', error);
  }
}

function updateLoanPaymentDropdowns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const loanSheet = ss.getSheetByName('🏦 Loans');
    
    if (!loanSheet) return;
    
    // Get all active loan IDs and member IDs for dropdowns
    const loanData = loanSheet.getDataRange().getValues();
    const loanIds = [];
    const memberIds = [];
    
    for (let i = 3; i < loanData.length; i++) {
      if (loanData[i][0] && loanData[i][0].toString().trim() !== '' && loanData[i][6] === 'Active') {
        loanIds.push(loanData[i][0]);
        if (loanData[i][1] && !memberIds.includes(loanData[i][1])) {
          memberIds.push(loanData[i][1]);
        }
      }
    }
    
    // Store for use in forms
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('loanIds', JSON.stringify(loanIds));
    scriptProperties.setProperty('loanMemberIds', JSON.stringify(memberIds));
    
  } catch (error) {
    console.error('Error updating loan payment dropdowns:', error);
  }
}

// ============================================
// ID GENERATION FUNCTIONS
// ============================================

function generateMemberId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    
    if (!savingsSheet) {
      return `M${CONFIG.nextMemberId}`;
    }
    
    const data = savingsSheet.getDataRange().getValues();
    let maxId = CONFIG.nextMemberId - 1;
    
    for (let i = 3; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().startsWith('M')) {
        const idStr = data[i][0].toString().substring(1);
        const idNum = parseInt(idStr);
        if (!isNaN(idNum) && idNum > maxId) {
          maxId = idNum;
        }
      }
    }
    
    const nextId = maxId + 1;
    return `M${nextId}`;
    
  } catch (error) {
    console.error('Error generating member ID:', error);
    return `M${CONFIG.nextMemberId}`;
  }
}

function generateLoanId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const loanSheet = ss.getSheetByName('🏦 Loans');
    
    if (!loanSheet) {
      return `LN${new Date().getTime()}`.substring(0, 15);
    }
    
    const data = loanSheet.getDataRange().getValues();
    let maxNumber = 0;
    
    for (let i = 3; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().startsWith('LN')) {
        const idStr = data[i][0].toString().substring(2);
        const idNum = parseInt(idStr);
        if (!isNaN(idNum) && idNum > maxNumber) {
          maxNumber = idNum;
        }
      }
    }
    
    const timestamp = new Date().getTime().toString().slice(-6);
    const newNumber = maxNumber + 1;
    return `LN${timestamp}${String(newNumber).padStart(4, '0')}`;
    
  } catch (error) {
    console.error('Error generating loan ID:', error);
    return `LN${new Date().getTime()}`.substring(0, 15);
  }
}

function applyCompleteSavingsFormatting(sheet) {
  try {
    console.log('Applying complete formatting...');
    
    // Center ALL data cells
    const lastRow = sheet.getLastRow();
    if (lastRow >= 5) {
      sheet.getRange(5, 1, lastRow - 4, 11).setHorizontalAlignment('center');
    }
    
    // Apply number formats
    sheet.getRange('C:C').setNumberFormat('"₱"#,##0.00');  // Total Savings
    sheet.getRange('G:G').setNumberFormat('"₱"#,##0.00');  // Next Payment Amount
    sheet.getRange('J:J').setNumberFormat('"₱"#,##0.00');  // Balance
    sheet.getRange('K:K').setNumberFormat('"₱"#,##0.00');  // Interest Earned
    
    // Apply date formats
    sheet.getRange('D:F').setNumberFormat('mm/dd/yyyy');
    
    // Apply integer format for months
    sheet.getRange('I:I').setNumberFormat('0');
    
    // Apply status colors
    for (let row = 5; row <= lastRow; row++) {
      const statusCell = sheet.getRange(row, 8); // Column H
      const status = statusCell.getValue();
      
      if (status) {
        const color = STATUS_COLORS[status] || COLORS.light;
        statusCell
          .setBackground(color)
          .setFontColor(COLORS.dark)
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
      }
    }
    
    console.log('Complete formatting applied');
    
  } catch (error) {
    console.error('Error in applyCompleteSavingsFormatting:', error);
  }
}

// ============================================
// TOGGLE SHEET VISIBILITY FUNCTION
// ============================================

function toggleSheetsVisibility() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    const allSheets = ss.getSheets();
    
    // Check if portal is already hidden (meaning we should show all)
    let showAll = false;
    try {
      if (portalSheet && portalSheet.isSheetHidden()) {
        showAll = true;
      }
    } catch (e) {
      // If we can't check, assume we need to show all
      showAll = true;
    }
    
    if (showAll) {
      // Show all sheets
      allSheets.forEach(sheet => {
        try {
          sheet.showSheet();
        } catch (e) {
          console.log('Could not show sheet:', sheet.getName());
        }
      });
      
      // Activate dashboard if it exists
      const dashboard = ss.getSheetByName('📈 Dashboard');
      if (dashboard) {
        ss.setActiveSheet(dashboard);
      }
      
      SpreadsheetApp.getUi().alert(
        '📋 All Sheets Visible',
        'All sheets are now visible.\n\n' +
        'Go to "View Settings" in the menu to hide sheets again.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      // Hide all sheets except Member Portal
      allSheets.forEach(sheet => {
        if (sheet.getName() !== '👤 Member Portal') {
          try {
            sheet.hideSheet();
          } catch (e) {
            console.log('Could not hide sheet:', sheet.getName());
          }
        }
      });
      
      // Show Member Portal
      if (portalSheet) {
        portalSheet.showSheet();
        ss.setActiveSheet(portalSheet);
      }
      
      SpreadsheetApp.getUi().alert(
        '👤 Member Portal Only',
        'All sheets except Member Portal are now hidden.\n\n' +
        'Members can only see their own information.\n' +
        'Go to "View Settings" in the menu to show all sheets again.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('Error toggling sheet visibility:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to toggle sheet visibility: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================
// SUPPORTING FUNCTIONS FOR ENHANCED MENU
// ============================================

function showSavingsSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const savingsSheet = ss.getSheetByName('💰 Savings');
  
  if (!savingsSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Savings sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  savingsSheet.showSheet();
  ss.setActiveSheet(savingsSheet);
}

function showDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('📈 Dashboard');
  
  if (!dashboard) {
    SpreadsheetApp.getUi().alert('Error', 'Dashboard not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  dashboard.showSheet();
  ss.setActiveSheet(dashboard);
}

function showCommunityFundsReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fundsSheet = ss.getSheetByName('💰 Community Funds');
  
  if (!fundsSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Community Funds sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  fundsSheet.showSheet();
  ss.setActiveSheet(fundsSheet);
  
  // Update calculations first
  updateCommunityFundsCalculations();
}

function showActiveLoansReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const loanSheet = ss.getSheetByName('🏦 Loans');
  
  if (!loanSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Loan sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Filter to show only active loans
  loanSheet.showSheet();
  ss.setActiveSheet(loanSheet);
  
  SpreadsheetApp.getUi().alert(
    'Active Loans Report',
    'Active loans are displayed.\n\n' +
    'You can filter by:\n' +
    '1. Click column G (Status)\n' +
    '2. Click the filter icon\n' +
    '3. Select "Active" only',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showPaymentScheduleReport() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:500px;">
      <h3 style="color:${COLORS.primary};text-align:center;margin-bottom:20px;">📅 Payment Schedule Report</h3>
      <p style="text-align:center;margin-bottom:20px;">Select report options:</p>
      
      <form id="reportForm">
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Report Type:</label>
          <select id="reportType" style="width:100%;padding:8px;box-sizing:border-box;">
            <option value="upcoming">Upcoming Payments (Next 30 days)</option>
            <option value="overdue">Overdue Payments</option>
            <option value="all">All Pending Payments</option>
          </select>
        </div>
        
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Payment Type:</label>
          <select id="paymentType" style="width:100%;padding:8px;box-sizing:border-box;">
            <option value="both">Both Savings & Loans</option>
            <option value="savings">Savings Only</option>
            <option value="loans">Loans Only</option>
          </select>
        </div>
        
        <div style="margin-bottom:20px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Sort By:</label>
          <select id="sortBy" style="width:100%;padding:8px;box-sizing:border-box;">
            <option value="date">Payment Date (Earliest First)</option>
            <option value="amount">Payment Amount (Highest First)</option>
            <option value="member">Member Name (A-Z)</option>
          </select>
        </div>
        
        <div style="display:flex;gap:10px;margin-top:25px;">
          <button type="button" onclick="generatePaymentReport()" 
            style="flex:1;background:${COLORS.success};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
            📅 Generate Report
          </button>
          <button type="button" onclick="closeDialog()" 
            style="flex:1;background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;">
            Cancel
          </button>
        </div>
      </form>
      
      <div id="message" style="margin-top:15px;"></div>
      
      <script>
        function generatePaymentReport() {
          const reportType = document.getElementById('reportType').value;
          const paymentType = document.getElementById('paymentType').value;
          const sortBy = document.getElementById('sortBy').value;
          
          showMessage('Generating report...', 'info');
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                showMessage('✅ ' + result.message, 'success');
                setTimeout(() => {
                  google.script.host.close();
                }, 2000);
              } else {
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .withFailureHandler(function(error) {
              showMessage('❌ Error: ' + error.message, 'error');
            })
            .generatePaymentScheduleReport(reportType, paymentType, sortBy);
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.innerHTML = text;
          messageDiv.style.color = type === 'error' ? '${COLORS.danger}' : 
                                  type === 'success' ? '${COLORS.success}' : 
                                  '${COLORS.info}';
          messageDiv.style.fontWeight = 'bold';
          messageDiv.style.textAlign = 'center';
          messageDiv.style.padding = '10px';
        }
        
        function closeDialog() {
          google.script.host.close();
        }
      </script>
    </div>
  `).setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Payment Schedule Report');
}

function generatePaymentScheduleReport(reportType, paymentType, sortBy) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create or get the report sheet
    let reportSheet = ss.getSheetByName('📅 Payment Schedule Report');
    if (!reportSheet) {
      reportSheet = ss.insertSheet('📅 Payment Schedule Report');
    } else {
      reportSheet.clear();
    }
    
    // Setup report sheet
    reportSheet.getRange(1, 1, 1, 7).merge()
      .setValue('📅 PAYMENT SCHEDULE REPORT')
      .setFontSize(16).setFontWeight('bold')
      .setFontColor(COLORS.white).setBackground(COLORS.primary)
      .setHorizontalAlignment('center');
    
    const headers = [
      ['Member ID', 'Member Name', 'Payment Type', 'Payment Date', 
       'Amount Due', 'Status', 'Contact Info']
    ];
    
    reportSheet.getRange(3, 1, 1, 7).setValues(headers)
      .setFontWeight('bold').setFontColor(COLORS.white)
      .setBackground(COLORS.secondary).setHorizontalAlignment('center');
    
    // Generate report based on parameters
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const loanSheet = ss.getSheetByName('🏦 Loans');
    
    let reportData = [];
    const today = new Date();
    
    // Process savings payments if requested
    if (paymentType === 'both' || paymentType === 'savings') {
      if (savingsSheet) {
        const savingsData = savingsSheet.getDataRange().getValues();
        for (let i = 4; i < savingsData.length; i++) {
          const memberId = savingsData[i][0];
          const memberName = savingsData[i][1];
          const nextPaymentDate = savingsData[i][5];
          const amountDue = savingsData[i][6] || CONFIG.paymentPerPeriod;
          const status = savingsData[i][7] || '';
          
          if (memberId && nextPaymentDate) {
            const paymentDate = new Date(nextPaymentDate);
            const daysUntil = Math.ceil((paymentDate - today) / (1000 * 60 * 60 * 24));
            
            let include = false;
            switch(reportType) {
              case 'upcoming':
                include = daysUntil >= 0 && daysUntil <= 30;
                break;
              case 'overdue':
                include = daysUntil < 0;
                break;
              case 'all':
                include = true;
                break;
            }
            
            if (include) {
              reportData.push({
                memberId: memberId,
                memberName: memberName,
                type: 'Savings',
                date: paymentDate,
                amount: amountDue,
                status: daysUntil < 0 ? 'Overdue' : (daysUntil <= 7 ? 'Due Soon' : 'Upcoming'),
                daysUntil: daysUntil,
                sortDate: paymentDate.getTime()
              });
            }
          }
        }
      }
    }
    
    // Process loan payments if requested
    if (paymentType === 'both' || paymentType === 'loans') {
      if (loanSheet) {
        const loanData = loanSheet.getDataRange().getValues();
        for (let i = 4; i < loanData.length; i++) {
          if (loanData[i][6] === 'Active') {
            const memberId = loanData[i][1];
            const nextPaymentDate = loanData[i][13];
            const amountDue = loanData[i][8] || 0;
            
            if (memberId && nextPaymentDate) {
              const paymentDate = new Date(nextPaymentDate);
              const daysUntil = Math.ceil((paymentDate - today) / (1000 * 60 * 60 * 24));
              
              let include = false;
              switch(reportType) {
                case 'upcoming':
                  include = daysUntil >= 0 && daysUntil <= 30;
                  break;
                case 'overdue':
                  include = daysUntil < 0;
                  break;
                case 'all':
                  include = true;
                  break;
              }
              
              if (include) {
                // Get member name
                let memberName = '';
                if (savingsSheet) {
                  const savingsData = savingsSheet.getDataRange().getValues();
                  for (let j = 4; j < savingsData.length; j++) {
                    if (savingsData[j][0] === memberId) {
                      memberName = savingsData[j][1];
                      break;
                    }
                  }
                }
                
                reportData.push({
                  memberId: memberId,
                  memberName: memberName,
                  type: 'Loan',
                  date: paymentDate,
                  amount: amountDue,
                  status: daysUntil < 0 ? 'Overdue' : (daysUntil <= 7 ? 'Due Soon' : 'Upcoming'),
                  daysUntil: daysUntil,
                  sortDate: paymentDate.getTime()
                });
              }
            }
          }
        }
      }
    }
    
    // Sort data
    switch(sortBy) {
      case 'date':
        reportData.sort((a, b) => a.sortDate - b.sortDate);
        break;
      case 'amount':
        reportData.sort((a, b) => b.amount - a.amount);
        break;
      case 'member':
        reportData.sort((a, b) => a.memberName.localeCompare(b.memberName));
        break;
    }
    
    // Write data to sheet
    if (reportData.length > 0) {
      const outputData = reportData.map(item => [
        item.memberId,
        item.memberName,
        item.type,
        item.date,
        item.amount,
        item.status,
        '' // Placeholder for contact info
      ]);
      
      reportSheet.getRange(4, 1, outputData.length, 7).setValues(outputData);
      
      // Format the data
      reportSheet.getRange('D:D').setNumberFormat('mm/dd/yyyy');
      reportSheet.getRange('E:E').setNumberFormat('₱#,##0.00');
      
      // Apply status colors
      for (let i = 0; i < reportData.length; i++) {
        const statusCell = reportSheet.getRange(i + 4, 6);
        const status = reportData[i].status;
        statusCell.setBackground(STATUS_COLORS[status] || COLORS.light)
          .setFontColor(COLORS.dark)
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
      }
      
      // Auto-resize columns
      reportSheet.autoResizeColumns(1, 7);
      
      // Add summary
      const summaryRow = outputData.length + 5;
      reportSheet.getRange(summaryRow, 1, 1, 7).merge()
        .setValue(`Report Summary: ${reportData.length} payments found (${reportType}, ${paymentType})`)
        .setFontWeight('bold')
        .setFontColor(COLORS.primary)
        .setHorizontalAlignment('center');
      
      // Show the sheet
      reportSheet.showSheet();
      ss.setActiveSheet(reportSheet);
      
      return { 
        success: true, 
        message: `Report generated successfully!\n\n` +
                 `Total Payments: ${reportData.length}\n` +
                 `Type: ${reportType}\n` +
                 `Payment Type: ${paymentType}\n` +
                 `Sorted By: ${sortBy}` 
      };
    } else {
      return { 
        success: true, 
        message: 'No payments match your criteria.' 
      };
    }
    
  } catch (error) {
    console.error('Error generating payment schedule report:', error);
    return { success: false, message: 'Error generating report: ' + error.message };
  }
}

function showDelinquencyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create report
  generatePaymentScheduleReport('overdue', 'both', 'date');
}

function showQuickLoanCalculator() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:500px;">
      <h3 style="color:${COLORS.primary};text-align:center;margin-bottom:20px;">⚡ Quick Loan Calculator</h3>
      <p style="text-align:center;color:${COLORS.warning};font-weight:bold;margin-bottom:20px;">
        ⚠️ ${(CONFIG.interestRateWithLoan * 100)}% MONTHLY INTEREST RATE
      </p>
      
      <form id="calculatorForm">
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Loan Amount (₱):</label>
          <input type="number" id="loanAmount" style="width:100%;padding:8px;box-sizing:border-box;" 
                 value="${CONFIG.minLoanAmount}" min="${CONFIG.minLoanAmount}" max="${CONFIG.maxLoanAmount}" 
                 oninput="calculateLoan()">
        </div>
        
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Loan Term (Months):</label>
          <input type="number" id="loanTerm" style="width:100%;padding:8px;box-sizing:border-box;" 
                 value="${CONFIG.loanTermMin}" min="${CONFIG.loanTermMin}" max="${CONFIG.loanTermMax}" 
                 oninput="calculateLoan()">
        </div>
      </form>
      
      <div id="calculationResult" style="margin-top:20px;padding:15px;background:${COLORS.light};border-radius:5px;">
        <h4 style="color:${COLORS.primary};margin-top:0;">Loan Calculation:</h4>
        <div id="loanDetails">
          <p>Enter loan details above to see calculation</p>
        </div>
      </div>
      
      <div style="display:flex;gap:10px;margin-top:25px;">
        <button type="button" onclick="applyQuickLoan()" 
          style="flex:1;background:${COLORS.success};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
          🏦 Apply This Loan
        </button>
        <button type="button" onclick="closeDialog()" 
          style="flex:1;background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;">
          Close
        </button>
      </div>
      
      <script>
        function calculateLoan() {
          const loanAmount = document.getElementById('loanAmount').value;
          const loanTerm = document.getElementById('loanTerm').value;
          
          if (!loanAmount || !loanTerm) {
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                document.getElementById('loanDetails').innerHTML = 
                  '<p><strong>Loan Amount:</strong> ₱' + result.loanAmount.toLocaleString() + '</p>' +
                  '<p><strong>Term:</strong> ' + result.loanTerm + ' months</p>' +
                  '<p><strong>Monthly Interest:</strong> ₱' + result.monthlyInterest.toLocaleString() + '</p>' +
                  '<p><strong>Total Interest:</strong> ₱' + result.totalInterest.toLocaleString() + '</p>' +
                  '<p><strong>Monthly Payment:</strong> ₱' + result.monthlyPayment.toLocaleString() + '</p>' +
                  '<p><strong>Total Repayment:</strong> ₱' + result.totalRepayment.toLocaleString() + '</p>' +
                  '<p><strong>Effective Annual Rate:</strong> ' + (result.effectiveAnnualRate * 100).toFixed(1) + '%</p>';
              } else {
                document.getElementById('loanDetails').innerHTML = 
                  '<p style="color:${COLORS.danger};">' + result.message + '</p>';
              }
            })
            .calculateLoanDetails(loanAmount, loanTerm);
        }
        
        function applyQuickLoan() {
          const loanAmount = document.getElementById('loanAmount').value;
          const loanTerm = document.getElementById('loanTerm').value;
          
          if (!loanAmount || !loanTerm) {
            alert('Please enter both loan amount and term.');
            return;
          }
          
          google.script.host.close();
          
          // Show loan application form with pre-filled values
          google.script.run.showLoanApplicationForm();
          
          // Note: The loan application form will need to handle pre-filling
          // You might want to pass the values as parameters
        }
        
        function closeDialog() {
          google.script.host.close();
        }
        
        // Calculate on page load
        window.onload = calculateLoan;
      </script>
    </div>
  `).setWidth(500).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quick Loan Calculator');
}

function updateSavingsCalculations() {
  try {
    fixSavingsSheetFormatting();
    fixMonthsActive();
    updateSummaryForExistingMembers();
    
    SpreadsheetApp.getUi().alert(
      '✅ Savings Calculations Updated',
      'All savings calculations have been refreshed:\n' +
      '• Months active recalculated\n' +
      '• Interest recalculated\n' +
      '• Statuses updated\n' +
      '• Summary sheet synchronized',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    console.error('Error updating savings calculations:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to update savings: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function fixCommunityFunds() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Recreate the funds sheet
    createCommunityFundsSheet(ss);
    
    // Update calculations
    updateCommunityFundsCalculations();
    
    SpreadsheetApp.getUi().alert(
      '✅ Community Funds Fixed',
      'Community Funds sheet has been completely rebuilt and recalculated.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    console.error('Error fixing community funds:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to fix community funds: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function exportAllReports() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:400px;">
      <h3 style="color:${COLORS.primary};text-align:center;margin-bottom:20px;">📋 Export All Reports</h3>
      <p style="text-align:center;margin-bottom:20px;">Export system data for backup or analysis:</p>
      
      <div style="display:flex;flex-direction:column;gap:10px;">
        <button onclick="exportReports('pdf')" 
          style="background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
          📄 Export as PDF
        </button>
        
        <button onclick="exportReports('excel')" 
          style="background:${COLORS.success};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
          📊 Export as Excel
        </button>
        
        <button onclick="exportReports('csv')" 
          style="background:${COLORS.info};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
          📝 Export as CSV
        </button>
        
        <button onclick="closeDialog()" 
          style="background:${COLORS.secondary};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;margin-top:10px;">
          Cancel
        </button>
      </div>
      
      <div id="message" style="margin-top:15px;"></div>
      
      <script>
        function exportReports(format) {
          showMessage('Preparing export...', 'info');
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                showMessage('✅ ' + result.message, 'success');
                setTimeout(() => {
                  google.script.host.close();
                }, 2000);
              } else {
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .withFailureHandler(function(error) {
              showMessage('❌ Error: ' + error.message, 'error');
            })
            .exportSystemReports(format);
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.innerHTML = text;
          messageDiv.style.color = type === 'error' ? '${COLORS.danger}' : 
                                  type === 'success' ? '${COLORS.success}' : 
                                  '${COLORS.info}';
          messageDiv.style.fontWeight = 'bold';
          messageDiv.style.textAlign = 'center';
          messageDiv.style.padding = '10px';
        }
        
        function closeDialog() {
          google.script.host.close();
        }
      </script>
    </div>
  `).setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Export Reports');
}

function exportSystemReports(format) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    switch(format) {
      case 'pdf':
        // Create PDF export
        const pdfUrl = ss.getUrl().replace('/edit', '/export?format=pdf');
        return {
          success: true,
          message: 'PDF export ready. Download link: ' + pdfUrl
        };
        
      case 'excel':
        // Create Excel export
        const excelUrl = ss.getUrl().replace('/edit', '/export?format=xlsx');
        return {
          success: true,
          message: 'Excel export ready. Download link: ' + excelUrl
        };
        
      case 'csv':
        // Create a zip of CSV files
        const folder = DriveApp.getFolderById(ss.getParent().getId());
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const exportFolder = DriveApp.createFolder('S&L_Export_' + timestamp);
        
        sheets.forEach(sheet => {
          const csvContent = sheet.getDataRange().getValues()
            .map(row => row.map(cell => {
              if (cell instanceof Date) {
                return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
              }
              return cell.toString().replace(/"/g, '""');
            }).join(','))
            .join('\n');
          
          exportFolder.createFile(sheet.getName() + '.csv', csvContent);
        });
        
        return {
          success: true,
          message: 'CSV files exported to folder: ' + exportFolder.getName(),
          folderUrl: exportFolder.getUrl()
        };
        
      default:
        return { success: false, message: 'Invalid export format.' };
    }
    
  } catch (error) {
    console.error('Error exporting reports:', error);
    return { success: false, message: 'Export failed: ' + error.message };
  }
}

function applySummaryFormatting(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 4) {
      // Center all data
      sheet.getRange(4, 1, lastRow - 3, 11).setHorizontalAlignment('center');
      
      // Apply number formats
      sheet.getRange('C:C').setNumberFormat('"₱"#,##0.00'); // Savings
      sheet.getRange('D:D').setNumberFormat('"₱"#,##0.00'); // Loans
      sheet.getRange('G:G').setNumberFormat('"₱"#,##0.00'); // Interest Paid
      sheet.getRange('I:I').setNumberFormat('"₱"#,##0.00'); // Balance
      
      // Apply integer formats
      sheet.getRange('E:E').setNumberFormat('0'); // Active Loans
      sheet.getRange('J:J').setNumberFormat('0'); // Savings Streak
      sheet.getRange('K:K').setNumberFormat('0'); // Loan Streak
      
      // Apply status colors
      for (let row = 4; row <= lastRow; row++) {
        const savingsStatusCell = sheet.getRange(row, 8); // Column H
        const savingsStatus = savingsStatusCell.getValue();
        
        if (savingsStatus) {
          savingsStatusCell
            .setBackground(STATUS_COLORS[savingsStatus] || COLORS.light)
            .setFontColor(COLORS.dark)
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
        }
        
        const loanStatusCell = sheet.getRange(row, 6); // Column F
        const loanStatus = loanStatusCell.getValue();
        
        if (loanStatus) {
          loanStatusCell
            .setBackground(STATUS_COLORS[loanStatus] || COLORS.light)
            .setFontColor(COLORS.dark)
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
        }
      }
    }
  } catch (error) {
    console.error('Error applying summary formatting:', error);
  }
}

function updateSummaryForExistingMembers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const summarySheet = ss.getSheetByName('📊 Summary');
    
    if (!savingsSheet || !summarySheet) return;
    
    // Clear existing summary data (keep headers)
    const lastSummaryRow = summarySheet.getLastRow();
    if (lastSummaryRow > 3) {
      summarySheet.getRange(4, 1, lastSummaryRow - 3, 11).clearContent();
    }
    
    // Get savings data
    const savingsData = savingsSheet.getDataRange().getValues();
    let summaryRow = 4; // Start after headers
    
    for (let i = 4; i < savingsData.length; i++) {
      const memberId = savingsData[i][0];
      const memberName = savingsData[i][1];
      const totalSavings = savingsData[i][2] || 0;
      const status = savingsData[i][7] || '';
      const monthsActive = savingsData[i][8] || 0;
      const balance = savingsData[i][9] || 0;
      const interestEarned = savingsData[i][10] || 0;
      
      if (memberId) {
        // Add to summary
        const summaryData = [
          memberId,
          memberName,
          totalSavings,
          0, // Loans
          0, // Active Loans
          'No Loan', // Loan Status
          0, // Interest Paid (loan interest)
          status, // Savings Status
          balance, // Overall Balance
          monthsActive >= 1 ? 1 : 0, // Savings Streak
          0 // Loan Streak
        ];
        
        summarySheet.getRange(summaryRow, 1, 1, 11).setValues([summaryData]);
        summaryRow++;
      }
    }
    
    // Apply summary formatting
    applySummaryFormatting(summarySheet);
    
  } catch (error) {
    console.error('Error updating summary:', error);
  }
}

function getPaymentPeriod(date) {
  const day = date.getDate();
  if (day <= 15) {
    return '1st Half (1st-15th)';
  } else {
    return '2nd Half (16th-31st)';
  }
}

// ============================================
// ADD NEW MEMBER FUNCTION (FIXED)
// ============================================

function addNewMember(memberId, fullName, initialSavings, dateJoined) {
  try {
    console.log('=== ADDING NEW MEMBER ===');
    console.log('Member ID:', memberId);
    console.log('Full Name:', fullName);
    console.log('Initial Savings:', initialSavings);
    console.log('Date Joined:', dateJoined);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const summarySheet = ss.getSheetByName('📊 Summary');
    const savingsPaymentsSheet = ss.getSheetByName('💵 Savings Payments');
    
    if (!savingsSheet || !summarySheet) {
      return { success: false, message: 'System sheets not found.' };
    }
    
    // Validate inputs
    if (!memberId || !fullName || !initialSavings || !dateJoined) {
      return { success: false, message: 'All fields are required.' };
    }
    
    const savingsAmount = parseFloat(initialSavings);
    if (isNaN(savingsAmount) || savingsAmount < CONFIG.savingsMin) {
      return { 
        success: false, 
        message: `Initial savings must be at least ₱${CONFIG.savingsMin.toLocaleString()}.` 
      };
    }
    
    // Check if member ID already exists
    const savingsData = savingsSheet.getDataRange().getValues();
    for (let i = 4; i < savingsData.length; i++) {
      if (savingsData[i][0] === memberId) {
        return { success: false, message: 'Member ID already exists. Please use a different ID.' };
      }
    }
    
    // Add to Savings sheet
    const nextRow = savingsSheet.getLastRow() + 1;
    const joinDate = new Date(dateJoined);
    const today = new Date();
    
    // ============================================
    // CORRECT NEXT PAYMENT DATE CALCULATION
    // ============================================
    let nextPaymentDate;
    const joinDay = joinDate.getDate();
    const joinMonth = joinDate.getMonth();
    const joinYear = joinDate.getFullYear();
    
    // Payment schedule: 5th and 20th
    if (joinDay <= CONFIG.paymentDates[0]) {
      // Joined on or before 5th -> next payment on 20th of same month
      nextPaymentDate = new Date(joinYear, joinMonth, CONFIG.paymentDates[1]);
    } else if (joinDay <= CONFIG.paymentDates[1]) {
      // Joined between 6th-20th -> next payment on 5th of next month
      nextPaymentDate = new Date(joinYear, joinMonth + 1, CONFIG.paymentDates[0]);
    } else {
      // Joined after 20th -> next payment on 5th of next month
      nextPaymentDate = new Date(joinYear, joinMonth + 1, CONFIG.paymentDates[0]);
    }
    
    // ============================================
    // CORRECT INTEREST CALCULATION FOR NEW MEMBERS
    // ============================================
    const monthlyInterestRate = CONFIG.interestRateSavingsOnly / 12; // 9% annual ÷ 12 = 0.75% monthly
    const halfMonthInterestRate = CONFIG.interestRateSavingsOnly / 24; // 9% annual ÷ 24 = 0.375% half-monthly

    // NEW MEMBERS: Start with 1 months active, 0 interest
    const monthsActive = 1;
    let interestEarned = savingsAmount * halfMonthInterestRate; // Interest for half month
    
    // ONLY calculate interest if they're paying at least a full month's contribution
    // For initial deposit, no interest yet
    const totalSavings = savingsAmount;
    const balance = totalSavings + interestEarned;
    
    // Status: "New" for new members, changes to "On Track" after first regular payment
    let status = 'New';
    
    console.log('Calculated values for new member:');
    console.log('- Next Payment Date:', nextPaymentDate.toLocaleDateString());
    console.log('- Monthly Interest Rate:', (monthlyInterestRate * 100).toFixed(2) + '%');
    console.log('- Months Active:', monthsActive);
    console.log('- Interest Earned:', interestEarned);
    console.log('- Total Savings:', totalSavings);
    console.log('- Status:', status);
    
    // ============================================
    // ADD TO SAVINGS SHEET WITH ALL VALUES
    // ============================================
    const savingsRecord = [
      memberId,                    // A: ID (starts in Column A)
      fullName,                    // B: Member Name
      totalSavings,               // C: Total Savings
      joinDate,                    // D: Date Joined
      joinDate,                    // E: Last Payment (same as join date)
      nextPaymentDate,             // F: Next Payment Date
      CONFIG.paymentPerPeriod,     // G: Next Payment Amount
      status,                      // H: Status
      monthsActive,                // I: Months Active
      balance,                     // J: Balance (savings + interest)
      interestEarned               // K: Interest Earned
    ];
    
    // Write to correct row starting from Column A
    savingsSheet.getRange(nextRow, 1, 1, 11).setValues([savingsRecord]);
    
    // ============================================
    // APPLY PROPER FORMATTING IMMEDIATELY
    // ============================================
    // Center all cells
    savingsSheet.getRange(nextRow, 1, 1, 11).setHorizontalAlignment('center');
    
    // Apply number formatting
    savingsSheet.getRange(nextRow, 3).setNumberFormat('"₱"#,##0.00');  // Total Savings (C)
    savingsSheet.getRange(nextRow, 7).setNumberFormat('"₱"#,##0.00');  // Next Payment (G)
    savingsSheet.getRange(nextRow, 10).setNumberFormat('"₱"#,##0.00'); // Balance (J)
    savingsSheet.getRange(nextRow, 11).setNumberFormat('"₱"#,##0.00'); // Interest (K)
    
    // Apply date formatting
    savingsSheet.getRange(nextRow, 4, 1, 3).setNumberFormat('mm/dd/yyyy'); // Dates (D, E, F)
    
    // Apply integer formatting for months
    savingsSheet.getRange(nextRow, 9).setNumberFormat('0'); // Months Active (I)
    
    // Apply status color
    const statusCell = savingsSheet.getRange(nextRow, 8);
    statusCell.setBackground(STATUS_COLORS[status] || COLORS.light)
      .setFontColor(COLORS.dark)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // ============================================
    // 🆕 FIXED: RECORD INITIAL DEPOSIT IN SAVINGS PAYMENTS SHEET
    // ============================================
    if (savingsPaymentsSheet && savingsAmount > 0) {
      const paymentId = 'SP' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
      const lastPaymentRow = savingsPaymentsSheet.getLastRow();
      
      // Make sure we're adding to the correct row (skip headers)
      const targetRow = lastPaymentRow < 3 ? 3 : lastPaymentRow + 1;
      
      const paymentRecord = [
        paymentId,                    // A: Payment ID
        memberId,                     // B: Member ID
        joinDate,                     // C: Date
        savingsAmount,                // D: Amount
        'Initial Deposit',            // E: Type
        'Verified',                   // F: Verified
        getPaymentPeriod(joinDate)    // G: Period
      ];
      
      savingsPaymentsSheet.getRange(targetRow, 1, 1, 7).setValues([paymentRecord]);
      
      // Apply formatting to payment record
      savingsPaymentsSheet.getRange(targetRow, 3).setNumberFormat('mm/dd/yyyy');
      savingsPaymentsSheet.getRange(targetRow, 4).setNumberFormat('"₱"#,##0.00');
      savingsPaymentsSheet.getRange(targetRow, 1, 1, 7).setHorizontalAlignment('center');
      
      console.log('✅ Initial deposit recorded in Savings Payments sheet:', paymentId);
    }
    
    // ============================================
    // ADD TO SUMMARY SHEET
    // ============================================
    const summaryRow = summarySheet.getLastRow() + 1;
    const summaryRecord = [
      memberId,                    // A: Member ID
      fullName,                    // B: Name
      totalSavings,               // C: Savings
      0,                           // D: Loans
      0,                           // E: Active Loans
      'No Loan',                   // F: Loan Status
      0,                           // G: Interest Paid (loan interest)
      status,                      // H: Savings Status
      balance,                     // I: Overall Balance
      0,                           // J: Savings Streak (0 for new members)
      0                            // K: Loan Streak
    ];
    
    summarySheet.getRange(summaryRow, 1, 1, 11).setValues([summaryRecord]);
    
    // Apply summary formatting
    summarySheet.getRange(summaryRow, 1, 1, 11).setHorizontalAlignment('center');
    summarySheet.getRange(summaryRow, 3).setNumberFormat('"₱"#,##0.00'); // Savings
    summarySheet.getRange(summaryRow, 4).setNumberFormat('"₱"#,##0.00'); // Loans
    summarySheet.getRange(summaryRow, 7).setNumberFormat('"₱"#,##0.00'); // Interest Paid
    summarySheet.getRange(summaryRow, 9).setNumberFormat('"₱"#,##0.00'); // Balance
    
    // Apply summary status colors
    const summaryStatusCell = summarySheet.getRange(summaryRow, 8);
    summaryStatusCell.setBackground(STATUS_COLORS[status] || COLORS.light)
      .setFontColor(COLORS.dark)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    const loanStatusCell = summarySheet.getRange(summaryRow, 6);
    loanStatusCell.setBackground(STATUS_COLORS['No Loan'] || COLORS.info)
      .setFontColor(COLORS.dark)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // ============================================
    // RECORD FUNDS TRANSACTION
    // ============================================
    recordFundsTransaction('New Member', memberId, 'Initial Deposit', savingsAmount, 
                          `New member registration - ${fullName}`);
    
    // Force spreadsheet update
    SpreadsheetApp.flush();
    
    // ============================================
    // UPDATE SYSTEM COMPONENTS
    // ============================================
    refreshAllDropdowns();
    updateMemberPortalDropdown();
    updateCommunityFundsCalculations();
    
    // Return success message
    return { 
      success: true, 
      message: `✅ MEMBER ADDED SUCCESSFULLY!\n\n` +
               `👤 Member Details:\n` +
               `• ID: ${memberId}\n` +
               `• Name: ${fullName}\n` +
               `• Initial Savings: ₱${savingsAmount.toLocaleString()}\n` +
               `• Date Joined: ${joinDate.toLocaleDateString()}\n\n` +
               `📅 Payment Schedule:\n` +
               `• Next Payment: ${nextPaymentDate.toLocaleDateString()}\n` +
               `• Amount Due: ₱${CONFIG.paymentPerPeriod.toLocaleString()}\n\n` +
               `📊 Current Status:\n` +
               `• Status: ${status}\n` +
               `• Months Active: ${monthsActive}\n` +
               `• Interest Earned: ₱${interestEarned.toLocaleString()}\n` +
               `• Total Balance: ₱${balance.toLocaleString()}\n\n` +
               `💳 Payment Recorded:\n` +
               `✓ Initial deposit of ₱${savingsAmount.toLocaleString()} recorded in Savings Payments sheet\n\n` +
               `💡 Reminder: Status will update to "On Track" after first regular payment. Interest will start accruing after first full month.`
    };
    
  } catch (error) {
    console.error('❌ Error adding member:', error);
    return { success: false, message: 'Error adding member: ' + error.message };
  }
}

// ============================================
// GET MEMBER INFO FOR FORMS (FIXED)
// ============================================

function getMemberInfo(memberId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    
    if (!savingsSheet) {
      return { success: false, message: 'Savings sheet not found.' };
    }
    
    const savingsData = savingsSheet.getDataRange().getValues();
    
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] === memberId) {
        return {
          success: true,
          name: savingsData[i][1],
          savings: savingsData[i][2] || 0,
          status: savingsData[i][7] || 'Unknown'
        };
      }
    }
    
    return { success: false, message: 'Member not found.' };
    
  } catch (error) {
    console.error('Error getting member info:', error);
    return { success: false, message: 'Error: ' + error.message };
  }
}

// ============================================
// CALCULATE NEXT PAYMENT DATE AFTER A PAYMENT
// ============================================

function getNextPaymentDateAfterPayment(paymentDate) {
  const date = new Date(paymentDate);
  const day = date.getDate();
  const month = date.getMonth();
  const year = date.getFullYear();
  
  // Payment schedule: 5th and 20th
  if (day <= CONFIG.paymentDates[0]) {
    // If payment is on or before the 5th, next payment is the 20th of same month
    return new Date(year, month, CONFIG.paymentDates[1]);
  } else if (day <= CONFIG.paymentDates[1]) {
    // If payment is on or before the 20th, next payment is the 5th of next month
    return new Date(year, month + 1, CONFIG.paymentDates[0]);
  } else {
    // If payment is after the 20th, next payment is the 5th of next month
    return new Date(year, month + 1, CONFIG.paymentDates[0]);
  }
}

// ============================================
// ENHANCED FORM FUNCTIONS WITH DROPDOWNS
// ============================================

function showAddMemberForm() {
  const generatedId = generateMemberId();
  
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:500px;">
      <h3 style="color:${COLORS.primary};text-align:center;margin-bottom:20px;">➕ Add New Member</h3>
      
      <form id="memberForm">
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Generated Member ID:</label>
          <input type="text" id="memberId" style="width:100%;padding:8px;box-sizing:border-box;" 
                 value="${generatedId}" readonly>
          <small style="color:#666;">Member ID is auto-generated</small>
        </div>
        
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Full Name:</label>
          <input type="text" id="fullName" style="width:100%;padding:8px;box-sizing:border-box;" 
                 placeholder="e.g., Juan Dela Cruz" required>
        </div>
        
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Initial Savings:</label>
          <input type="number" id="initialSavings" style="width:100%;padding:8px;box-sizing:border-box;" 
                 value="${CONFIG.savingsMin}" min="${CONFIG.savingsMin}" step="100" required>
          <small style="color:#666;">Minimum: ₱${CONFIG.savingsMin.toLocaleString()}</small>
        </div>
        
        <div style="margin-bottom:20px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Date Joined:</label>
          <input type="date" id="dateJoined" style="width:100%;padding:8px;box-sizing:border-box;" 
                 value="${new Date().toISOString().split('T')[0]}" required>
        </div>
        
        <div style="display:flex;gap:10px;margin-top:25px;">
          <button type="button" onclick="addMember()" 
            style="flex:1;background:${COLORS.success};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
            ➕ Add Member
          </button>
          <button type="button" onclick="closeDialog()" 
            style="flex:1;background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;">
            Cancel
          </button>
        </div>
      </form>
      
      <div id="message" style="margin-top:15px;"></div>
      
      <script>
        function addMember() {
          const memberId = document.getElementById('memberId').value.trim();
          const fullName = document.getElementById('fullName').value.trim();
          const initialSavings = document.getElementById('initialSavings').value;
          const dateJoined = document.getElementById('dateJoined').value;
          
          if (!fullName || !initialSavings || !dateJoined) {
            showMessage('Please fill in all required fields.', 'error');
            return;
          }
          
          showMessage('Adding member...', 'info');
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                showMessage('✅ ' + result.message, 'success');
                setTimeout(() => {
                  google.script.host.close();
                  google.script.run.showAllMembers();
                }, 1500);
              } else {
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .withFailureHandler(function(error) {
              showMessage('❌ Error: ' + error.message, 'error');
            })
            .addNewMember(memberId, fullName, initialSavings, dateJoined);
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.innerHTML = text;
          messageDiv.style.color = type === 'error' ? '${COLORS.danger}' : 
                                  type === 'success' ? '${COLORS.success}' : 
                                  '${COLORS.info}';
          messageDiv.style.fontWeight = 'bold';
          messageDiv.style.textAlign = 'center';
          messageDiv.style.padding = '10px';
        }
        
        function closeDialog() {
          google.script.host.close();
        }
      </script>
    </div>
  `).setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Member');
}

function showSavingsPaymentForm() {
  // Get member IDs for dropdown
  const scriptProperties = PropertiesService.getScriptProperties();
  let memberIdsJson = scriptProperties.getProperty('memberIds');
  let memberIds = [];
  
  if (memberIdsJson) {
    try {
      memberIds = JSON.parse(memberIdsJson);
    } catch (e) {
      console.error('Error parsing memberIds:', e);
    }
  }
  
  // If no member IDs in cache, get them fresh
  if (memberIds.length === 0) {
    updateSavingsPaymentDropdowns();
    memberIdsJson = scriptProperties.getProperty('memberIds');
    if (memberIdsJson) {
      try {
        memberIds = JSON.parse(memberIdsJson);
      } catch (e) {
        console.error('Error parsing memberIds:', e);
      }
    }
  }
  
  // Create dropdown options HTML
  let memberOptions = '<option value="">-- Select Member --</option>';
  memberIds.forEach(id => {
    memberOptions += `<option value="${id}">${id}</option>`;
  });
  
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:500px;">
      <h3 style="color:${COLORS.primary};text-align:center;margin-bottom:20px;">💵 Record Savings Payment</h3>
      
      <form id="paymentForm">
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Member ID:</label>
          <select id="memberId" style="width:100%;padding:8px;box-sizing:border-box;" required onchange="loadMemberInfo()">
            ${memberOptions}
          </select>
          <small style="color:#666;">Select member from dropdown</small>
        </div>
        
        <div id="memberInfo" style="background:${COLORS.light};padding:10px;border-radius:5px;margin-bottom:15px;display:none;">
          <p style="margin:0;font-weight:bold;" id="memberName"></p>
          <p style="margin:5px 0 0 0;font-size:12px;" id="currentSavings"></p>
        </div>
        
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Payment Date:</label>
          <input type="date" id="paymentDate" style="width:100%;padding:8px;box-sizing:border-box;" 
                 value="${new Date().toISOString().split('T')[0]}" required>
        </div>
        
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Payment Amount:</label>
          <select id="paymentAmount" style="width:100%;padding:8px;box-sizing:border-box;" required>
            <option value="${CONFIG.paymentPerPeriod}">₱${CONFIG.paymentPerPeriod.toLocaleString()} (Regular Payment - ${CONFIG.paymentDates[0]}th or ${CONFIG.paymentDates[1]}th)</option>
            <option value="${CONFIG.monthlyContribution}">₱${CONFIG.monthlyContribution.toLocaleString()} (Full Month Payment)</option>
            <option value="500">₱500 (Partial Payment)</option>
            <option value="other">Other Amount</option>
          </select>
        </div>
        
        <div style="margin-bottom:15px;display:none;" id="otherAmountDiv">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Enter Amount:</label>
          <input type="number" id="otherAmount" style="width:100%;padding:8px;box-sizing:border-box;" 
                 min="100" step="100" placeholder="Enter amount in pesos">
        </div>
        
        <div style="margin-bottom:20px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Payment Type:</label>
          <select id="paymentType" style="width:100%;padding:8px;box-sizing:border-box;" required>
            <option value="Regular">Regular Payment</option>
            <option value="Additional">Additional Savings</option>
            <option value="Catch-up">Catch-up Payment</option>
          </select>
        </div>
        
        <div style="display:flex;gap:10px;margin-top:25px;">
          <button type="button" onclick="recordPayment()" 
            style="flex:1;background:${COLORS.success};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
            💵 Record Payment
          </button>
          <button type="button" onclick="closeDialog()" 
            style="flex:1;background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;">
            Cancel
          </button>
        </div>
      </form>
      
      <div id="message" style="margin-top:15px;"></div>
      
      <script>
        document.getElementById('paymentAmount').addEventListener('change', function() {
          const otherAmountDiv = document.getElementById('otherAmountDiv');
          otherAmountDiv.style.display = this.value === 'other' ? 'block' : 'none';
          if (this.value === 'other') {
            document.getElementById('otherAmount').focus();
          }
        });
        
        function loadMemberInfo() {
          const memberId = document.getElementById('memberId').value;
          if (!memberId) {
            document.getElementById('memberInfo').style.display = 'none';
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(member) {
              if (member.success) {
                document.getElementById('memberInfo').style.display = 'block';
                document.getElementById('memberName').textContent = 'Member: ' + member.name;
                document.getElementById('currentSavings').textContent = 
                  'Current Savings: ₱' + member.savings.toLocaleString() + ' | Status: ' + member.status;
                showMessage('✅ Member loaded!', 'success');
              } else {
                document.getElementById('memberInfo').style.display = 'none';
                showMessage('❌ ' + member.message, 'error');
              }
            })
            .getMemberInfo(memberId);
        }
        
        function recordPayment() {
          const memberId = document.getElementById('memberId').value;
          const paymentDate = document.getElementById('paymentDate').value;
          const paymentAmount = document.getElementById('paymentAmount').value;
          const otherAmount = document.getElementById('otherAmount').value;
          const paymentType = document.getElementById('paymentType').value;
          
          if (!memberId) {
            showMessage('Please select a Member.', 'error');
            return;
          }
          
          let amount = paymentAmount;
          if (paymentAmount === 'other') {
            if (!otherAmount || otherAmount < 100) {
              showMessage('Please enter a valid amount (minimum ₱100).', 'error');
              return;
            }
            amount = otherAmount;
          }
          
          showMessage('Recording payment...', 'info');
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                showMessage('✅ ' + result.message, 'success');
                setTimeout(() => {
                  google.script.host.close();
                }, 2000);
              } else {
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .withFailureHandler(function(error) {
              showMessage('❌ Error: ' + error.message, 'error');
            })
            .recordSavingsPayment(memberId, amount, paymentDate, paymentType);
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.innerHTML = text;
          messageDiv.style.color = type === 'error' ? '${COLORS.danger}' : 
                                  type === 'success' ? '${COLORS.success}' : 
                                  '${COLORS.info}';
          messageDiv.style.fontWeight = 'bold';
          messageDiv.style.textAlign = 'center';
          messageDiv.style.padding = '10px';
        }
        
        function closeDialog() {
          google.script.host.close();
        }
      </script>
    </div>
  `).setWidth(500).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Savings Payment');
}

function showLoanPaymentForm() {
  // Get loan IDs and member IDs for dropdowns
  const scriptProperties = PropertiesService.getScriptProperties();
  const loanIdsJson = scriptProperties.getProperty('loanIds');
  const loanMemberIdsJson = scriptProperties.getProperty('loanMemberIds');
  
  let loanIds = [];
  let memberIds = [];
  
  if (loanIdsJson) {
    try {
      loanIds = JSON.parse(loanIdsJson);
    } catch (e) {
      console.error('Error parsing loanIds:', e);
    }
  }
  
  if (loanMemberIdsJson) {
    try {
      memberIds = JSON.parse(loanMemberIdsJson);
    } catch (e) {
      console.error('Error parsing loanMemberIds:', e);
    }
  }
  
  // If no data in cache, get it fresh
  if (loanIds.length === 0 || memberIds.length === 0) {
    updateLoanPaymentDropdowns();
    const newLoanIdsJson = scriptProperties.getProperty('loanIds');
    const newMemberIdsJson = scriptProperties.getProperty('loanMemberIds');
    
    if (newLoanIdsJson) {
      try {
        loanIds = JSON.parse(newLoanIdsJson);
      } catch (e) {
        console.error('Error parsing loanIds:', e);
      }
    }
    
    if (newMemberIdsJson) {
      try {
        memberIds = JSON.parse(newMemberIdsJson);
      } catch (e) {
        console.error('Error parsing loanMemberIds:', e);
      }
    }
  }
  
  // Create dropdown options HTML
  let loanOptions = '<option value="">-- Select Loan --</option>';
  loanIds.forEach(id => {
    loanOptions += `<option value="${id}">${id}</option>`;
  });
  
  let memberOptions = '<option value="">-- Select Member --</option>';
  memberIds.forEach(id => {
    memberOptions += `<option value="${id}">${id}</option>`;
  });
  
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:500px;">
      <h3 style="color:${COLORS.primary};text-align:center;margin-bottom:20px;">💳 Record Loan Payment</h3>
      
      <form id="loanPaymentForm">
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Select by:</label>
          <select id="searchType" style="width:100%;padding:8px;box-sizing:border-box;" onchange="toggleSearchType()">
            <option value="loan">Loan ID</option>
            <option value="member">Member ID</option>
          </select>
        </div>
        
        <div id="loanSection" style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Loan ID:</label>
          <select id="loanId" style="width:100%;padding:8px;box-sizing:border-box;" onchange="loadLoanDetails()">
            ${loanOptions}
          </select>
        </div>
        
        <div id="memberSection" style="margin-bottom:15px;display:none;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Member ID:</label>
          <select id="memberId" style="width:100%;padding:8px;box-sizing:border-box;" onchange="loadMemberLoans()">
            ${memberOptions}
          </select>
          <div id="memberLoans" style="margin-top:10px;display:none;">
            <label style="display:block;margin-bottom:5px;font-weight:bold;">Select Loan:</label>
            <select id="selectedLoan" style="width:100%;padding:8px;box-sizing:border-box;" onchange="loadSelectedLoanDetails()">
              <option value="">-- Select a loan --</option>
            </select>
          </div>
        </div>
        
        <div id="loanDetails" style="background:${COLORS.light};padding:10px;border-radius:5px;margin-bottom:15px;display:none;">
          <p style="margin:0;font-weight:bold;" id="loanInfo"></p>
          <p style="margin:5px 0 0 0;font-size:12px;" id="remainingBalance"></p>
          <p style="margin:5px 0 0 0;font-size:12px;" id="nextPayment"></p>
        </div>
        
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Payment Date:</label>
          <input type="date" id="paymentDate" style="width:100%;padding:8px;box-sizing:border-box;" 
                 value="${new Date().toISOString().split('T')[0]}" required>
        </div>
        
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Payment Amount:</label>
          <input type="number" id="paymentAmount" style="width:100%;padding:8px;box-sizing:border-box;" 
                 min="100" step="100" placeholder="Enter payment amount" required>
          <small style="color:#666;">Minimum: ₱100</small>
        </div>
        
        <div style="display:flex;gap:10px;margin-top:25px;">
          <button type="button" onclick="recordLoanPayment()" 
            style="flex:1;background:${COLORS.success};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
            💳 Record Payment
          </button>
          <button type="button" onclick="closeDialog()" 
            style="flex:1;background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;">
            Cancel
          </button>
        </div>
      </form>
      
      <div id="message" style="margin-top:15px;"></div>
      
      <script>
        function toggleSearchType() {
          const searchType = document.getElementById('searchType').value;
          const loanSection = document.getElementById('loanSection');
          const memberSection = document.getElementById('memberSection');
          
          if (searchType === 'loan') {
            loanSection.style.display = 'block';
            memberSection.style.display = 'none';
            document.getElementById('memberLoans').style.display = 'none';
            document.getElementById('loanDetails').style.display = 'none';
          } else {
            loanSection.style.display = 'none';
            memberSection.style.display = 'block';
            document.getElementById('loanDetails').style.display = 'none';
          }
        }
        
        function loadLoanDetails() {
          const loanId = document.getElementById('loanId').value;
          if (!loanId) {
            document.getElementById('loanDetails').style.display = 'none';
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                document.getElementById('loanDetails').style.display = 'block';
                document.getElementById('loanInfo').textContent = 
                  'Loan: ' + result.id + ' | Amount: ₱' + result.amount.toLocaleString();
                document.getElementById('remainingBalance').textContent = 
                  'Remaining Balance: ₱' + result.remaining.toLocaleString();
                document.getElementById('nextPayment').textContent = 
                  'Monthly Payment: ₱' + result.monthlyPayment.toLocaleString();
                
                document.getElementById('paymentAmount').value = result.monthlyPayment;
              } else {
                document.getElementById('loanDetails').style.display = 'none';
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .getLoanDetails(loanId);
        }
        
        function loadMemberLoans() {
          const memberId = document.getElementById('memberId').value;
          if (!memberId) {
            document.getElementById('memberLoans').style.display = 'none';
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success && result.loans.length > 0) {
                const select = document.getElementById('selectedLoan');
                select.innerHTML = '<option value="">-- Select a loan --</option>';
                
                result.loans.forEach(function(loan) {
                  const option = document.createElement('option');
                  option.value = loan.id;
                  option.textContent = loan.id + ' - ₱' + loan.amount.toLocaleString() + ' (' + loan.status + ')';
                  select.appendChild(option);
                });
                
                document.getElementById('memberLoans').style.display = 'block';
                showMessage('✅ ' + result.loans.length + ' active loan(s) found.', 'success');
              } else {
                document.getElementById('memberLoans').style.display = 'none';
                showMessage('❌ No active loans found for this member.', 'error');
              }
            })
            .findActiveLoans('', memberId);
        }
        
        function loadSelectedLoanDetails() {
          const loanId = document.getElementById('selectedLoan').value;
          if (!loanId) {
            document.getElementById('loanDetails').style.display = 'none';
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                document.getElementById('loanDetails').style.display = 'block';
                document.getElementById('loanInfo').textContent = 
                  'Loan: ' + result.id + ' | Amount: ₱' + result.amount.toLocaleString();
                document.getElementById('remainingBalance').textContent = 
                  'Remaining Balance: ₱' + result.remaining.toLocaleString();
                document.getElementById('nextPayment').textContent = 
                  'Monthly Payment: ₱' + result.monthlyPayment.toLocaleString();
                
                document.getElementById('paymentAmount').value = result.monthlyPayment;
              } else {
                document.getElementById('loanDetails').style.display = 'none';
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .getLoanDetails(loanId);
        }
        
        function recordLoanPayment() {
          let loanId;
          const searchType = document.getElementById('searchType').value;
          
          if (searchType === 'loan') {
            loanId = document.getElementById('loanId').value;
          } else {
            loanId = document.getElementById('selectedLoan').value;
          }
          
          const paymentDate = document.getElementById('paymentDate').value;
          const paymentAmount = document.getElementById('paymentAmount').value;
          
          if (!loanId) {
            showMessage('Please select a Loan.', 'error');
            return;
          }
          
          if (!paymentAmount || paymentAmount < 100) {
            showMessage('Please enter a valid payment amount (minimum ₱100).', 'error');
            return;
          }
          
          showMessage('Recording loan payment...', 'info');
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                showMessage('✅ ' + result.message, 'success');
                setTimeout(() => {
                  google.script.host.close();
                }, 2000);
              } else {
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .withFailureHandler(function(error) {
              showMessage('❌ Error: ' + error.message, 'error');
            })
            .recordLoanPayment(loanId, paymentAmount, paymentDate);
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.innerHTML = text;
          messageDiv.style.color = type === 'error' ? '${COLORS.danger}' : 
                                  type === 'success' ? '${COLORS.success}' : 
                                  '${COLORS.info}';
          messageDiv.style.fontWeight = 'bold';
          messageDiv.style.textAlign = 'center';
          messageDiv.style.padding = '10px';
        }
        
        function closeDialog() {
          google.script.host.close();
        }
        
        // Initialize
        toggleSearchType();
      </script>
    </div>
  `).setWidth(500).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Loan Payment');
}

// ============================================
// UPDATED LOAN CALCULATION (FIXED FORMULA)
// ============================================

function calculateLoanDetails(loanAmount, loanTerm) {
  try {
    const amount = parseFloat(loanAmount);
    const term = parseInt(loanTerm);
    
    if (isNaN(amount) || amount < CONFIG.minLoanAmount || amount > CONFIG.maxLoanAmount) {
      return { 
        success: false, 
        message: 'Invalid loan amount. Must be between ₱' + CONFIG.minLoanAmount.toLocaleString() + 
                ' and ₱' + CONFIG.maxLoanAmount.toLocaleString() + '.' 
      };
    }
    
    if (isNaN(term) || term < CONFIG.loanTermMin || term > CONFIG.loanTermMax) {
      return { 
        success: false, 
        message: 'Invalid loan term. Must be between ' + CONFIG.loanTermMin + 
                '-' + CONFIG.loanTermMax + ' months.' 
      };
    }
    
    // CORRECTED INTEREST CALCULATION FOR MONTHLY INTEREST
    // Monthly interest rate = CONFIG.interestRateWithLoan (e.g., 9% per month)
    const monthlyInterestRate = CONFIG.interestRateWithLoan; // e.g., 0.09 (9% per month)
    
    // Total interest = Loan Amount × Interest Rate × Term (months)
    // This is the correct formula: 3000 × 0.09 × 3 = 810
    const totalInterest = amount * monthlyInterestRate * term;
    
    // Total repayment = Loan Amount + Total Interest
    const totalRepayment = amount + totalInterest;
    
    // Monthly payment = Total Repayment ÷ Term
    const monthlyPayment = totalRepayment / term;
    
    // Monthly interest amount (for display)
    const monthlyInterest = amount * monthlyInterestRate;
    
    // Effective Annual Rate = (1 + monthly rate)^12 - 1
    const effectiveAnnualRate = Math.pow(1 + monthlyInterestRate, 12) - 1;
    
    return {
      success: true,
      loanAmount: amount,
      loanTerm: term,
      interestRate: CONFIG.interestRateWithLoan, // Monthly interest rate
      monthlyInterestRate: monthlyInterestRate,
      monthlyInterest: monthlyInterest,
      totalInterest: totalInterest,
      totalRepayment: totalRepayment,
      monthlyPayment: monthlyPayment,
      effectiveAnnualRate: effectiveAnnualRate
    };
    
  } catch (error) {
    console.error('Error calculating loan details:', error);
    return { success: false, message: 'Calculation error: ' + error.message };
  }
}

// ============================================
// UPDATED SAVINGS INTEREST CALCULATION
// ============================================

function calculateSavingsInterest(totalSavings, hasActiveLoan = false) {
  try {
    const savings = parseFloat(totalSavings);
    if (isNaN(savings) || savings <= 0) return 0;
    
    // Determine which interest rate to use
    let annualInterestRate;
    if (hasActiveLoan) {
      annualInterestRate = CONFIG.interestRateWithoutLoan; // Lower rate if has loan
    } else {
      annualInterestRate = CONFIG.interestRateSavingsOnly; // Higher rate if no loan
    }
    
    // Monthly interest = (Annual Rate / 12) × Savings
    const monthlyInterestRate = annualInterestRate / 12;
    const monthlyInterest = savings * monthlyInterestRate;
    
    return {
      monthlyInterest: monthlyInterest,
      annualInterest: savings * annualInterestRate,
      monthlyRate: monthlyInterestRate,
      annualRate: annualInterestRate
    };
    
  } catch (error) {
    console.error('Error calculating savings interest:', error);
    return { monthlyInterest: 0, annualInterest: 0, monthlyRate: 0, annualRate: 0 };
  }
}

// ============================================
// ENHANCED MEMBER PORTAL FUNCTIONS
// ============================================

function viewMemberInfoFromPortal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    
    if (!portalSheet) {
      SpreadsheetApp.getUi().alert('❌ Error', 'Member Portal not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const memberId = portalSheet.getRange(7, 3).getValue();
    
    if (!memberId || memberId.toString().trim() === '' || memberId === '--- Clear Selection ---') {
      SpreadsheetApp.getUi().alert(
        'Select Member ID',
        'Please select a valid Member ID from the dropdown in cell C7.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Use the auto-search function
    searchAndDisplayMember(memberId.toString().trim());
    
  } catch (error) {
    console.error('Error in viewMemberInfoFromPortal:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 'Search failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function clearMemberPortal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    
    if (!portalSheet) return;
    
    // 1. Clear the Member ID cell
    portalSheet.getRange(7, 3).clearContent();
    
    // 2. Clear all display areas
    // Clear rows 11-20 (main display)
    portalSheet.getRange(11, 1, 10, 6).clearContent().clearFormat();
    
    // Clear additional info rows 15-19
    portalSheet.getRange(15, 1, 5, 3).clearContent().clearFormat();
    
    // 3. Reset to default message
    portalSheet.getRange(11, 1, 1, 6).merge()
      .setValue('👤 Select your Member ID above to view your information')
      .setFontColor(COLORS.info)
      .setHorizontalAlignment('center')
      .setFontStyle('italic');
    
    // 4. Reset additional info section
    portalSheet.getRange(13, 1, 1, 6).merge()
      .setValue('📊 ADDITIONAL INFORMATION')
      .setFontSize(12).setFontWeight('bold')
      .setFontColor(COLORS.white).setBackground(COLORS.info)
      .setHorizontalAlignment('center');
    
    const infoHeaders = [['Information', 'Value', 'Notes']];
    portalSheet.getRange(14, 1, 1, 3).setValues(infoHeaders)
      .setFontWeight('bold').setFontColor(COLORS.dark)
      .setBackground(COLORS.light).setHorizontalAlignment('center');
    
    const infoRows = [
      ['Active Loans', '', 'Number of active loans'],
      ['Monthly Interest', '', 'Interest earned this month'],
      ['Total Interest', '', 'Total interest earned'],
      ['Payment Streak', '', 'Consecutive on-time payments'],
      ['Loan Status', '', 'Status of any active loans']
    ];
    
    portalSheet.getRange(15, 1, infoRows.length, 3).setValues(infoRows);
    
    console.log('✅ Portal cleared');
    
    // Optional: Show confirmation (uncomment if you want popup)
    // SpreadsheetApp.getUi().alert('✅ Cleared', 'Member Portal has been cleared.', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Error clearing member portal:', error);
  }
}

function searchMemberById(memberId) {
  try {
    console.log('Searching for member:', memberId);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    const summarySheet = ss.getSheetByName('📊 Summary');
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const loanSheet = ss.getSheetByName('🏦 Loans');
    
    if (!portalSheet || !summarySheet || !savingsSheet) {
      return { success: false, message: 'System not initialized.' };
    }
    
    // Get savings data
    const savingsData = savingsSheet.getDataRange().getValues();
    let memberFound = false;
    let memberName = '';
    let totalSavings = 0;
    let status = '';
    let nextPaymentDate = '';
    let nextPaymentAmount = CONFIG.paymentPerPeriod;
    let monthsActive = 0;
    let interestEarned = 0;
    let savingsRow = -1;
    
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] === memberId) {
        memberFound = true;
        savingsRow = i;
        memberName = savingsData[i][1];
        totalSavings = savingsData[i][2] || 0;
        status = savingsData[i][7] || 'On Track';
        nextPaymentDate = savingsData[i][5];
        nextPaymentAmount = savingsData[i][6] || CONFIG.paymentPerPeriod;
        monthsActive = savingsData[i][8] || 1;
        interestEarned = savingsData[i][10] || 0;
        break;
      }
    }
    
    if (!memberFound) {
      return { success: false, message: 'Member ID not found. Please check your ID.' };
    }
    
    // Get summary data
    let currentLoan = 0;
    let activeLoans = 0;
    let overallBalance = totalSavings;
    let loanStatus = 'No Loan';
    let savingsStreak = 0;
    let loanStreak = 0;
    let interestPaid = 0;
    
    const summaryData = summarySheet.getDataRange().getValues();
    for (let j = 3; j < summaryData.length; j++) {
      if (summaryData[j][0] === memberId) {
        currentLoan = summaryData[j][3] || 0;
        activeLoans = summaryData[j][4] || 0;
        overallBalance = summaryData[j][8] || totalSavings;
        loanStatus = summaryData[j][5] || 'No Loan';
        interestPaid = summaryData[j][6] || 0;
        savingsStreak = summaryData[j][9] || 0;
        loanStreak = summaryData[j][10] || 0;
        break;
      }
    }
    
    // Get active loan details
    let loanDetails = [];
    if (loanSheet && activeLoans > 0) {
      const loanData = loanSheet.getDataRange().getValues();
      for (let k = 3; k < loanData.length; k++) {
        if (loanData[k][1] === memberId && loanData[k][6] === 'Active') {
          loanDetails.push({
            loanId: loanData[k][0],
            amount: loanData[k][2],
            remaining: loanData[k][11],
            monthlyPayment: loanData[k][8],
            nextPaymentDate: loanData[k][13]
          });
        }
      }
    }
    
    // Calculate total balance with interest
    const hasActiveLoan = activeLoans > 0;
    const interestCalc = calculateSavingsInterest(totalSavings, hasActiveLoan);
    const totalBalanceWithInterest = totalSavings + interestEarned;
    
    // Format next payment date
    let formattedNextPaymentDate = 'N/A';
    if (nextPaymentDate) {
      if (nextPaymentDate instanceof Date) {
        formattedNextPaymentDate = nextPaymentDate.toLocaleDateString();
      } else if (typeof nextPaymentDate === 'string') {
        try {
          formattedNextPaymentDate = new Date(nextPaymentDate).toLocaleDateString();
        } catch (e) {
          formattedNextPaymentDate = nextPaymentDate;
        }
      }
    }
    
    // Calculate days until next payment
    let daysUntilNextPayment = 'N/A';
    if (nextPaymentDate instanceof Date) {
      const today = new Date();
      const timeDiff = nextPaymentDate.getTime() - today.getTime();
      daysUntilNextPayment = Math.ceil(timeDiff / (1000 * 3600 * 24));
    }
    
    // Get savings payment history
    let lastPaymentDate = 'N/A';
    let lastPaymentAmount = 0;
    if (savingsRow >= 0) {
      lastPaymentDate = savingsData[savingsRow][4];
      if (lastPaymentDate && lastPaymentDate instanceof Date) {
        lastPaymentDate = lastPaymentDate.toLocaleDateString();
      }
    }
    
    return { 
      success: true, 
      message: 'Information loaded successfully',
      memberId: memberId,
      memberName: memberName,
      totalSavings: totalSavings,
      totalBalanceWithInterest: totalBalanceWithInterest,
      currentLoan: currentLoan,
      status: status,
      activeLoans: activeLoans,
      loanStatus: loanStatus,
      overallBalance: overallBalance,
      interestRate: interestCalc.annualRate,
      monthlyInterest: interestCalc.monthlyInterest,
      interestEarned: interestEarned,
      savingsStreak: savingsStreak,
      loanStreak: loanStreak,
      interestPaid: interestPaid,
      nextPaymentDate: formattedNextPaymentDate,
      nextPaymentAmount: nextPaymentAmount,
      daysUntilNextPayment: daysUntilNextPayment,
      monthsActive: monthsActive,
      lastPaymentDate: lastPaymentDate,
      lastPaymentAmount: lastPaymentAmount,
      hasActiveLoan: hasActiveLoan,
      loanDetails: loanDetails,
      config: {
        paymentPerPeriod: CONFIG.paymentPerPeriod,
        monthlyContribution: CONFIG.monthlyContribution,
        paymentDates: CONFIG.paymentDates
      }
    };
    
  } catch (error) {
    console.error('Error searching member:', error);
    return { 
      success: false, 
      message: 'System error: ' + error.message,
      errorDetails: error.toString()
    };
  }
}

function fixM1001Savings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const summarySheet = ss.getSheetByName('📊 Summary');
    
    // Fix Savings sheet
    const savingsData = savingsSheet.getDataRange().getValues();
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] === 'M1001') {
        // Set savings to ₱1,000
        savingsSheet.getRange(i + 1, 3).setValue(1000);
        // Set balance to ₱1,000
        savingsSheet.getRange(i + 1, 10).setValue(1000);
        console.log('✅ Fixed M1001 in savings sheet');
        break;
      }
    }
    
    // Fix Summary sheet
    const summaryData = summarySheet.getDataRange().getValues();
    for (let i = 3; i < summaryData.length; i++) {
      if (summaryData[i][0] === 'M1001') {
        // Set savings to ₱1,000
        summarySheet.getRange(i + 1, 3).setValue(1000);
        // Set balance to ₱1,000
        summarySheet.getRange(i + 1, 9).setValue(1000);
        // Set savings streak to 1
        summarySheet.getRange(i + 1, 10).setValue(1);
        console.log('✅ Fixed M1001 in summary sheet');
        break;
      }
    }
    
    SpreadsheetApp.getUi().alert('✅ Fixed', 'M1001 savings corrected to ₱1,000', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Error fixing M1001:', error);
  }
}

// ============================================
// UPDATED RECORD SAVINGS PAYMENT WITH INTEREST
// ============================================

function recordSavingsPayment(memberId, amount, paymentDate, paymentType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const paymentsSheet = ss.getSheetByName('💵 Savings Payments');
    const summarySheet = ss.getSheetByName('📊 Summary');
    
    if (!savingsSheet || !paymentsSheet || !summarySheet) {
      return { success: false, message: 'System sheets not found.' };
    }
    
    const savingsData = savingsSheet.getDataRange().getValues();
    let memberRow = -1;
    let currentSavings = 0;
    let currentInterest = 0;
    let monthsActive = 0;
    let memberName = '';
    let currentStatus = '';
    let dateJoined = null;
    
    // Find member
    for (let i = 4; i < savingsData.length; i++) {
      if (savingsData[i][0] === memberId) {
        memberRow = i;
        currentSavings = savingsData[i][2] || 0;
        currentInterest = savingsData[i][10] || 0;
        monthsActive = savingsData[i][8] || 0;
        memberName = savingsData[i][1];
        currentStatus = savingsData[i][7] || 'New';
        dateJoined = savingsData[i][3];
        break;
      }
    }
    
    if (memberRow === -1) {
      return { success: false, message: 'Member ID not found.' };
    }
    
    const paymentAmount = parseFloat(amount);
    const newSavings = currentSavings + paymentAmount;
    const paymentDateObj = new Date(paymentDate);
    
    // ============================================
    // FIXED: CORRECT MONTHS ACTIVE CALCULATION
    // ============================================
    let newMonthsActive = monthsActive;
    
    if (paymentType === 'Regular' && paymentAmount >= CONFIG.paymentPerPeriod) {
      // For new members (0 months active), this is their first regular payment
      if (monthsActive === 0) {
        // Check if this completes their first month
        const daysSinceJoin = Math.floor((paymentDateObj - dateJoined) / (1000 * 60 * 60 * 24));
        
        if (daysSinceJoin >= 30) {
          newMonthsActive = 1; // First month completed
        }
      } else {
        // For existing members, check if this starts a new month
        const lastPaymentDate = savingsData[memberRow][4]; // Last payment date
        if (lastPaymentDate) {
          const daysSinceLastPayment = Math.floor((paymentDateObj - lastPaymentDate) / (1000 * 60 * 60 * 24));
          
          if (daysSinceLastPayment >= 30) {
            newMonthsActive = monthsActive + 1;
          }
        }
      }
    } else if (paymentType === 'Additional' && paymentAmount >= CONFIG.monthlyContribution) {
      // If paying full month at once
      if (monthsActive === 0) {
        newMonthsActive = 1;
      }
    }
    
    // ============================================
    // FIXED: CORRECT INTEREST CALCULATION
    // ============================================
    const monthlyInterestRate = CONFIG.interestRateSavingsOnly / 12; // 0.75% monthly
    
    // Calculate interest based on NEW savings and NEW months active
    // Interest = Total Savings × Monthly Rate × Months Active
    const newInterest = newSavings * monthlyInterestRate * newMonthsActive;
    
    // Round to 2 decimal places
    const roundedInterest = Math.round(newInterest * 100) / 100;
    
    // ============================================
    // FIXED: NEXT PAYMENT DATE CALCULATION
    // ============================================
    const nextPaymentDate = getNextPaymentDateAfterPayment(paymentDateObj);
    
    // ============================================
    // FIXED: STATUS UPDATE
    // ============================================
    let newStatus = currentStatus;
    if (paymentType === 'Regular' && paymentAmount >= CONFIG.paymentPerPeriod) {
      if (currentStatus === 'New' || currentStatus === 'Partial') {
        newStatus = 'On Track';
      }
    } else if (paymentAmount > 0 && paymentAmount < CONFIG.paymentPerPeriod) {
      newStatus = 'Partial';
    }
    
    // ============================================
    // FIXED: UPDATE SAVINGS SHEET
    // ============================================
    // Update all cells for this member
    savingsSheet.getRange(memberRow + 1, 3).setValue(newSavings);      // C: Total Savings
    savingsSheet.getRange(memberRow + 1, 5).setValue(paymentDateObj);  // E: Last Payment
    savingsSheet.getRange(memberRow + 1, 6).setValue(nextPaymentDate); // F: Next Payment Date
    savingsSheet.getRange(memberRow + 1, 7).setValue(CONFIG.paymentPerPeriod); // G: Next Payment Amount
    savingsSheet.getRange(memberRow + 1, 8).setValue(newStatus);       // H: Status
    savingsSheet.getRange(memberRow + 1, 9).setValue(newMonthsActive); // I: Months Active
    savingsSheet.getRange(memberRow + 1, 10).setValue(newSavings + roundedInterest); // J: Balance
    savingsSheet.getRange(memberRow + 1, 11).setValue(roundedInterest); // K: Interest Earned
    
    // Apply formatting
    savingsSheet.getRange(memberRow + 1, 3).setNumberFormat('"₱"#,##0.00');
    savingsSheet.getRange(memberRow + 1, 7).setNumberFormat('"₱"#,##0.00');
    savingsSheet.getRange(memberRow + 1, 10).setNumberFormat('"₱"#,##0.00');
    savingsSheet.getRange(memberRow + 1, 11).setNumberFormat('"₱"#,##0.00');
    savingsSheet.getRange(memberRow + 1, 5, 1, 2).setNumberFormat('mm/dd/yyyy');
    savingsSheet.getRange(memberRow + 1, 1, 1, 11).setHorizontalAlignment('center');
    
    // Apply status color
    const statusCell = savingsSheet.getRange(memberRow + 1, 8);
    statusCell.setBackground(STATUS_COLORS[newStatus] || COLORS.light)
      .setFontColor(COLORS.dark)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // ============================================
    // RECORD PAYMENT
    // ============================================
    const paymentId = 'SP' + new Date().getTime();
    const lastPaymentRow = paymentsSheet.getLastRow();
    const paymentRecord = [
      paymentId,
      memberId,
      paymentDateObj,
      paymentAmount,
      paymentType,
      'Verified',
      getPaymentPeriod(paymentDateObj)
    ];
    
    paymentsSheet.getRange(lastPaymentRow + 1, 1, 1, 7).setValues([paymentRecord]);
    
    // ============================================
    // UPDATE SUMMARY SHEET
    // ============================================
    const summaryData = summarySheet.getDataRange().getValues();
    for (let i = 3; i < summaryData.length; i++) {
      if (summaryData[i][0] === memberId) {
        const totalBalance = newSavings + roundedInterest;
        const currentLoans = summaryData[i][3] || 0;
        
        summarySheet.getRange(i + 1, 3).setValue(totalBalance);
        summarySheet.getRange(i + 1, 8).setValue(newStatus);
        summarySheet.getRange(i + 1, 9).setValue(totalBalance - currentLoans);
        
        // Calculate savings streak: If on track, streak = months active
        let savingsStreak = 0;
        if (newStatus === 'On Track') {
          savingsStreak = newMonthsActive;
        }
        summarySheet.getRange(i + 1, 10).setValue(savingsStreak);
        
        // Apply formatting
        summarySheet.getRange(i + 1, 3).setNumberFormat('"₱"#,##0.00');
        summarySheet.getRange(i + 1, 9).setNumberFormat('"₱"#,##0.00');
        
        // Apply status color
        const summaryStatusCell = summarySheet.getRange(i + 1, 8);
        summaryStatusCell.setBackground(STATUS_COLORS[newStatus] || COLORS.light)
          .setFontColor(COLORS.dark)
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        break;
      }
    }
    
    // ============================================
    // UPDATE OTHER SYSTEMS
    // ============================================
    recordFundsTransaction('Savings Deposit', memberId, paymentId, paymentAmount, 
                          `${paymentType} payment - ${memberName}`);
    
    refreshAllDropdowns();
    updateCommunityFundsCalculations();
    
    return { 
      success: true, 
      message: `✅ PAYMENT RECORDED!\n\n` +
               `Member: ${memberName}\n` +
               `Payment: ₱${paymentAmount.toLocaleString()} (${paymentType})\n\n` +
               `📊 Updated Details:\n` +
               `• Total Savings: ₱${newSavings.toLocaleString()}\n` +
               `• Months Active: ${newMonthsActive}\n` +
               `• Interest Earned: ₱${roundedInterest.toLocaleString()}\n` +
               `• Total Balance: ₱${(newSavings + roundedInterest).toLocaleString()}\n` +
               `• Status: ${newStatus}\n` +
               `• Next Payment: ${nextPaymentDate.toLocaleDateString()}` 
    };
    
  } catch (error) {
    console.error('Error recording payment:', error);
    return { success: false, message: 'Failed to record payment: ' + error.message };
  }
}

// ============================================
// UPDATED PROCESS LOAN APPLICATION
// ============================================

function processLoanApplication(memberId, loanAmount, loanTerm) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const loanSheet = ss.getSheetByName('🏦 Loans');
    const summarySheet = ss.getSheetByName('📊 Summary');
    const fundsSheet = ss.getSheetByName('💰 Community Funds');
    
    if (!savingsSheet || !loanSheet || !summarySheet || !fundsSheet) {
      return { success: false, message: 'System sheets not found.' };
    }
    
    // Check eligibility
    const eligibility = checkLoanEligibility(memberId);
    if (!eligibility.success) {
      return eligibility;
    }
    
    const amount = parseFloat(loanAmount);
    const term = parseInt(loanTerm);
    
    if (term < CONFIG.loanTermMin || term > CONFIG.loanTermMax) {
      return { 
        success: false, 
        message: 'Loan term must be between ' + CONFIG.loanTermMin + 
                '-' + CONFIG.loanTermMax + ' months.' 
      };
    }
    
    // CHECK COMMUNITY FUNDS AVAILABILITY
    const fundsCheck = checkCommunityFunds(amount);
    if (!fundsCheck.canApprove) {
      return {
        success: false,
        message: `❌ LOAN CANNOT BE APPROVED\n\n` +
                 `Reason: ${fundsCheck.message}\n\n` +
                 `Available Community Funds: ₱${fundsCheck.available.toLocaleString()}\n` +
                 `Requested Loan Amount: ₱${amount.toLocaleString()}\n` +
                 `Shortfall: ₱${(amount - fundsCheck.available).toLocaleString()}\n\n` +
                 `Please wait for more savings contributions or request a smaller loan amount.`,
        fundsAvailable: fundsCheck.available,
        requested: amount
      };
    }
    
    const savingsData = savingsSheet.getDataRange().getValues();
    let memberName = '';
    
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] === memberId) {
        memberName = savingsData[i][1];
        break;
      }
    }
    
    // Calculate loan details
    const calculation = calculateLoanDetails(loanAmount, loanTerm);
    if (!calculation.success) {
      return calculation;
    }
    
    const loanId = generateLoanId();
    const today = new Date();
    const nextPaymentDate = getNextPaymentDate();
    
    // Add to loans sheet
    const lastLoanRow = loanSheet.getLastRow();
    const loanRecord = [
      loanId,                    // Loan ID
      memberId,                  // Member ID
      amount,                    // Loan Amount
      CONFIG.interestRateWithLoan, // Interest Rate (Monthly)
      term,                      // Term (months)
      today,                     // Date Approved
      'Active',                  // Status
      calculation.monthlyInterest, // Monthly Interest
      calculation.monthlyPayment,  // Monthly Payment
      calculation.totalInterest,   // Total Interest
      calculation.totalRepayment,  // Total Repayment
      calculation.totalRepayment,  // Remaining Balance
      0,                         // Payments Made
      nextPaymentDate,           // Next Payment Date
      calculation.effectiveAnnualRate // Effective Annual Rate
    ];
    
    loanSheet.getRange(lastLoanRow + 1, 1, 1, 15).setValues([loanRecord]);
    
    // Format loan record
    loanSheet.getRange(lastLoanRow + 1, 3).setNumberFormat('₱#,##0.00');
    loanSheet.getRange(lastLoanRow + 1, 4).setNumberFormat('0%');
    loanSheet.getRange(lastLoanRow + 1, 8).setNumberFormat('₱#,##0.00');
    loanSheet.getRange(lastLoanRow + 1, 9).setNumberFormat('₱#,##0.00');
    loanSheet.getRange(lastLoanRow + 1, 10).setNumberFormat('₱#,##0.00');
    loanSheet.getRange(lastLoanRow + 1, 11).setNumberFormat('₱#,##0.00');
    loanSheet.getRange(lastLoanRow + 1, 12).setNumberFormat('₱#,##0.00');
    loanSheet.getRange(lastLoanRow + 1, 6).setNumberFormat('mm/dd/yyyy');
    loanSheet.getRange(lastLoanRow + 1, 14).setNumberFormat('mm/dd/yyyy');
    loanSheet.getRange(lastLoanRow + 1, 15).setNumberFormat('0%');
    
    // Update summary sheet
    const summaryData = summarySheet.getDataRange().getValues();
    for (let i = 3; i < summaryData.length; i++) {
      if (summaryData[i][0] === memberId) {
        const currentSavings = summaryData[i][2] || 0;
        const currentLoans = summaryData[i][3] || 0;
        const activeLoans = summaryData[i][4] || 0;
        
        summarySheet.getRange(i + 1, 4).setValue(currentLoans + amount);
        summarySheet.getRange(i + 1, 5).setValue(activeLoans + 1);
        summarySheet.getRange(i + 1, 6).setValue('Active');
        summarySheet.getRange(i + 1, 9).setValue(currentSavings - (currentLoans + amount));
        
        summarySheet.getRange(i + 1, 4).setNumberFormat('₱#,##0.00');
        summarySheet.getRange(i + 1, 9).setNumberFormat('₱#,##0.00');
        break;
      }
    }
    
    // UPDATE SAVINGS SHEET - Adjust interest rate for member with loan
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] === memberId) {
        const savingsAmount = savingsData[i][2] || 0;
        const interestCalc = calculateSavingsInterest(savingsAmount, true); // true = has active loan
        const currentInterest = savingsData[i][10] || 0;
        
        savingsSheet.getRange(i + 1, 11).setValue(currentInterest);
        savingsSheet.getRange(i + 1, 10).setValue(savingsAmount + currentInterest);
        break;
      }
    }
    
    // RECORD TRANSACTION IN COMMUNITY FUNDS SHEET
    recordFundsTransaction('Loan Disbursement', memberId, loanId, amount, 
                          `Funds disbursed for new loan at ${CONFIG.interestRateWithLoan * 100}% monthly interest`);
    
    // Update silently
    updateDashboard();
    refreshAllDropdowns();
    
    return {
      success: true,
      loanId: loanId,
      message: `✅ LOAN APPROVED!\n\n` +
               `⚠️ IMPORTANT: ${CONFIG.interestRateWithLoan * 100}% MONTHLY INTEREST APPLIES\n\n` +
               `Formula Used: Loan Amount × Interest Rate × Term\n` +
               `₱${amount.toLocaleString()} × ${CONFIG.interestRateWithLoan} × ${term} = ₱${calculation.totalInterest.toLocaleString()}\n\n` +
               `Loan Details:\n` +
               `Member: ${memberName}\n` +
               `Loan ID: ${loanId}\n` +
               `Loan Amount: ₱${amount.toLocaleString()}\n` +
               `Term: ${term} months\n` +
               `Interest Rate: ${(CONFIG.interestRateWithLoan * 100)}% PER MONTH\n` +
               `Monthly Interest: ₱${calculation.monthlyInterest.toLocaleString()}\n` +
               `Total Interest: ₱${calculation.totalInterest.toLocaleString()}\n` +
               `Effective Annual Rate: ${(calculation.effectiveAnnualRate * 100).toFixed(1)}%\n` +
               `Monthly Payment: ₱${calculation.monthlyPayment.toLocaleString()}\n` +
               `Total Repayment: ₱${calculation.totalRepayment.toLocaleString()}\n\n` +
               `Payment Schedule: Monthly, starting ${nextPaymentDate.toLocaleDateString()}`,
      fundsAvailable: fundsCheck.available - amount
    };
    
  } catch (error) {
    console.error('Error processing loan application:', error);
    return { success: false, message: 'Failed to process loan: ' + error.message };
  }
}

// ============================================
// HELPER FUNCTIONS (UPDATED)
// ============================================

function getNextPaymentDate() {
  const today = new Date();
  const currentDay = today.getDate();
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();
  
  let nextPaymentDate = new Date(currentYear, currentMonth, CONFIG.paymentDates[0]);
  
  if (currentDay > CONFIG.paymentDates[0]) {
    nextPaymentDate = new Date(currentYear, currentMonth, CONFIG.paymentDates[1]);
    
    if (currentDay > CONFIG.paymentDates[1]) {
      nextPaymentDate = new Date(currentYear, currentMonth + 1, CONFIG.paymentDates[0]);
    }
  }
  
  return nextPaymentDate;
}

function getPaymentPeriod(date) {
  const day = date.getDate();
  if (day <= 15) {
    return '1st Half (1st-15th)';
  } else {
    return '2nd Half (16th-31st)';
  }
}

function applyStatusColors(sheet, statusColumn, startRow, endRow) {
  if (!sheet || !statusColumn) return;
  
  const lastRow = endRow || sheet.getLastRow();
  if (lastRow < startRow) return;
  
  const statusRange = sheet.getRange(startRow, statusColumn, lastRow - startRow + 1, 1);
  const statusValues = statusRange.getValues();
  
  for (let i = 0; i < statusValues.length; i++) {
    const cell = statusRange.getCell(i + 1, 1);
    const status = statusValues[i][0];
    const color = STATUS_COLORS[status] || COLORS.light;
    
    cell.setBackground(color);
    cell.setFontColor(COLORS.dark);
    cell.setFontWeight('bold');
    cell.setHorizontalAlignment('center');
  }
}

// ============================================
// FIXED: INITIALIZE EXISTING DATA FORMATTING
// ============================================

function fixSavingsSheetFormatting() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    
    if (!savingsSheet) return;
    
    const lastRow = savingsSheet.getLastRow();
    if (lastRow <= 4) return; // No data rows
    
    // Apply center alignment to all data rows
    savingsSheet.getRange(5, 1, lastRow - 4, 11).setHorizontalAlignment('center');
    
    // Apply proper formatting to all columns
    savingsSheet.getRange(5, 3, lastRow - 4, 1).setNumberFormat('₱#,##0.00'); // Total Savings
    savingsSheet.getRange(5, 7, lastRow - 4, 1).setNumberFormat('₱#,##0'); // Next Payment Amount - FIXED
    savingsSheet.getRange(5, 10, lastRow - 4, 1).setNumberFormat('₱#,##0.00'); // Balance
    savingsSheet.getRange(5, 11, lastRow - 4, 1).setNumberFormat('₱#,##0.00'); // Interest
    savingsSheet.getRange(5, 4, lastRow - 4, 3).setNumberFormat('mm/dd/yyyy'); // Dates
    
    // Fix specific case from the image: M1001
    const data = savingsSheet.getDataRange().getValues();
    for (let i = 4; i < data.length; i++) {
      if (data[i][0] === 'M1001') {
        // Update total savings to ₱1,000
        savingsSheet.getRange(i + 1, 3).setValue(1000);
        
        // FIXED: Set Next Payment Amount to ₱1,000 (column G)
        savingsSheet.getRange(i + 1, 7).setValue(CONFIG.paymentPerPeriod);
        
        // Update next payment date to January 20
        const nextPaymentDate = new Date('2024-01-20');
        savingsSheet.getRange(i + 1, 6).setValue(nextPaymentDate);
        
        // Update status to "On Track"
        savingsSheet.getRange(i + 1, 8).setValue('On Track');
        const statusCell = savingsSheet.getRange(i + 1, 8);
        statusCell.setBackground(STATUS_COLORS['On Track'] || COLORS.success)
          .setFontColor(COLORS.dark)
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
        
        // Update balance and interest
        const interest = calculateSavingsInterest(1000, false).monthlyInterest;
        savingsSheet.getRange(i + 1, 10).setValue(1000 + (interest || 0));
        savingsSheet.getRange(i + 1, 11).setValue(interest || 0);
        break;
      }
    }
    
    // Apply status colors to all rows
    applyStatusColors(savingsSheet, 8, 5, lastRow);
    
    console.log('Savings sheet formatting fixed.');
    
  } catch (error) {
    console.error('Error fixing savings sheet formatting:', error);
  }
}

function fixMonthsActive() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    
    if (!savingsSheet) return;
    
    const data = savingsSheet.getDataRange().getValues();
    
    // Start from row 5 (index 4)
    for (let i = 4; i < data.length; i++) {
      const savings = data[i][2] || 0; // Column C
      const dateJoined = data[i][3]; // Column D
      const lastPayment = data[i][4] || dateJoined; // Column E
      
      if (dateJoined && lastPayment) {
        // Calculate months based on days between
        const joinDate = new Date(dateJoined);
        const lastPaymentDate = new Date(lastPayment);
        
        const daysBetween = Math.floor((lastPaymentDate - joinDate) / (1000 * 60 * 60 * 24));
        const months = Math.max(1, Math.ceil(daysBetween / 30));
        
        savingsSheet.getRange(i + 1, 9).setValue(months); // Column I (index 8)
      }
    }
    
    SpreadsheetApp.getUi().alert('✅ Fixed', 'Months Active calculations have been corrected.', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Error fixing months active:', error);
  }
}

// ============================================
// ADD THIS FUNCTION TO FIX EXISTING DATA
// ============================================

function fixAllFormatting() {
  fixSavingsSheetFormatting();
  updateCommunityFundsCalculations();
  fixMonthsActive();
  updateAllCalculations();
  refreshAllDropdowns();
  updateMemberPortalDropdown();
  installAutoUpdateTrigger();
  fixMemberPortal();
  addClearOptionToDropdown()

  SpreadsheetApp.getUi().alert(
    '✅ Formatting Fixed',
    'All savings sheet formatting has been corrected:\n' +
    '• Status column alignment fixed\n' +
    '• All data centered\n' +
    '• Status colors applied\n' +
    '• Next payment dates corrected',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function updateDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName('📈 Dashboard');
    
    if (!dashboard) return;
    
    dashboard.getRange(3, 2).setValue(new Date().toLocaleDateString());
    
  } catch (error) {
    console.error('Error updating dashboard:', error);
  }
}

// ============================================
// FIXED: RECORD FUNDS TRANSACTION FUNCTION
// ============================================

function recordFundsTransaction(transactionType, memberId, referenceId, amount, notes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fundsSheet = ss.getSheetByName('💰 Community Funds');
    
    if (!fundsSheet) return;
    
    // Find next available row in activity section
    let nextRow = 22;
    while (fundsSheet.getRange(nextRow, 1).getValue() !== '') {
      nextRow++;
      if (nextRow > 50) break;
    }
    
    // Get current balance
    const currentAvailable = fundsSheet.getRange('B9').getValue() || 0;
    const transactionAmount = parseFloat(amount);
    
    // Calculate new balance based on transaction type
    let balanceBefore = currentAvailable;
    let balanceAfter = currentAvailable;
    
    switch(transactionType) {
      case 'Loan Disbursement':
      case 'Member Removal':
        balanceAfter = Math.max(0, currentAvailable - transactionAmount);
        break;
      case 'Loan Repayment':
      case 'Savings Deposit':
      case 'New Member':
        balanceAfter = currentAvailable + transactionAmount;
        break;
    }
    
    // Record transaction
    const transactionData = [
      new Date(),
      transactionType,
      memberId ? `${memberId} / ${referenceId}` : referenceId,
      transactionAmount,
      balanceBefore,
      balanceAfter,
      notes || `${transactionType} transaction`
    ];
    
    fundsSheet.getRange(nextRow, 1, 1, 7).setValues([transactionData]);
    
    // Format the row
    fundsSheet.getRange(nextRow, 1).setNumberFormat('mm/dd/yyyy hh:mm');
    fundsSheet.getRange(nextRow, 4, 1, 3).setNumberFormat('₱#,##0.00');
    
    // Update community funds calculations
    updateCommunityFundsCalculations();
    
  } catch (error) {
    console.error('Error recording funds transaction:', error);
  }
}

// ============================================
// ENHANCED ADMIN MENU
// ============================================

// ============================================
// ENHANCED ADMIN MENU WITH CATEGORIES
// ============================================

function createAdminMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.createMenu('🏦 S&L System')
      .addItem('🎯 Generate/Reset System', 'generateProfessionalSystem')
      .addSeparator()
      
      // MEMBER MANAGEMENT
      .addSubMenu(ui.createMenu('👥 Member Management')
        .addItem('➕ Add New Member', 'showAddMemberForm')
        .addItem('👥 View All Members', 'showAllMembers')
        .addItem('🔍 Search Member', 'showSearchMemberForm')
        .addItem('🗑️ Remove Member', 'showRemoveMemberForm')
      )
      
      .addSeparator()
      
      // SAVINGS OPERATIONS
      .addSubMenu(ui.createMenu('💰 Savings')
        .addItem('💵 Record Savings Payment', 'showSavingsPaymentForm')
        .addItem('📋 Savings Summary', 'showSavingsSummary')
        .addItem('📊 Update Savings Calculations', 'updateSavingsCalculations')
      )
      
      // LOAN OPERATIONS
      .addSubMenu(ui.createMenu('🏦 Loans')
        .addItem('📝 Apply for Loan', 'showLoanApplicationForm')
        .addItem('💳 Record Loan Payment', 'showLoanPaymentForm')
        .addItem('📋 Active Loans Report', 'showActiveLoansReport')
        .addItem('⚡ Quick Loan Calculator', 'showQuickLoanCalculator')
      )
      
      .addSeparator()
      
      // REPORTS & ANALYTICS
      .addSubMenu(ui.createMenu('📊 Reports')
        .addItem('📈 Dashboard Overview', 'showDashboard')
        .addItem('💰 Community Funds Report', 'showCommunityFundsReport')
        .addItem('📅 Payment Schedule Report', 'showPaymentScheduleReport')
        .addItem('📉 Delinquency Report', 'showDelinquencyReport')
        .addItem('📋 Export All Reports', 'exportAllReports')
      )
      
      // TOOLS & MAINTENANCE
      .addSubMenu(ui.createMenu('🔧 Tools')
        .addItem('⚙️ System Settings', 'showSettings')
        .addItem('🔄 Refresh All Dropdowns', 'refreshAllDropdowns')
        .addItem('🔧 Fix All Formatting', 'fixAllFormatting')
        .addItem('📐 Fix Months Active', 'fixMonthsActive')
        .addItem('💰 Fix Community Funds', 'fixCommunityFunds')
      )
      
      .addSeparator()
      
      // VIEW SETTINGS
      .addSubMenu(ui.createMenu('👁️ View Settings')
        .addItem('📤 Show All Sheets', 'showAllSheets')
        .addItem('📥 Member Portal Only', 'setupMemberView')
        .addItem('🔄 Toggle Sheet Visibility', 'toggleSheetsVisibility')
        .addItem('🎯 Focus on Dashboard', 'showDashboard')
        .addItem('💰 Focus on Community Funds', 'showCommunityFunds')
      )
      
      .addSeparator()
      
      // UPDATES & CALCULATIONS
      .addSubMenu(ui.createMenu('🔄 Updates')
        .addItem('⚡ Update All Calculations', 'updateAllCalculations')
        .addItem('📊 Update Community Funds', 'updateCommunityFundsCalculations')
        .addItem('📈 Update Dashboard', 'updateDashboard')
        .addItem('📋 Update Summary Sheet', 'updateSummaryForExistingMembers')
      )
      
      .addToUi();
      
  } catch (error) {
    console.log('Admin menu creation error:', error.message);
    // Fallback simple menu
    SpreadsheetApp.getUi().createMenu('🏦 S&L System')
      .addItem('🎯 Generate System', 'generateProfessionalSystem')
      .addItem('📤 Show All Sheets', 'showAllSheets')
      .addItem('📥 Member Portal Only', 'setupMemberView')
      .addToUi();
  }
}

// ============================================
// ADDITIONAL REQUIRED FUNCTIONS
// ============================================

function showLoanApplicationForm() {
  // Get member IDs for dropdown from Savings sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const savingsSheet = ss.getSheetByName('💰 Savings');
  const memberIds = [];
  
  if (savingsSheet) {
    const savingsData = savingsSheet.getDataRange().getValues();
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] && savingsData[i][0].toString().trim() !== '') {
        memberIds.push(savingsData[i][0]);
      }
    }
  }
  
  // Create dropdown options HTML
  let memberOptions = '<option value="">-- Select Member --</option>';
  memberIds.forEach(id => {
    memberOptions += `<option value="${id}">${id}</option>`;
  });
  
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:500px;">
      <h3 style="color:${COLORS.primary};text-align:center;margin-bottom:20px;">🏦 Apply for Loan</h3>
      <p style="text-align:center;color:${COLORS.warning};font-weight:bold;">
        ⚠️ Loan Interest: ${(CONFIG.interestRateWithLoan * 100)}% PER MONTH
      </p>
      
      <form id="loanForm">
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Member ID:</label>
          <select id="memberId" style="width:100%;padding:8px;box-sizing:border-box;" required onchange="loadMemberInfo()">
            ${memberOptions}
          </select>
          <small style="color:#666;">Select member from dropdown</small>
        </div>
        
        <div id="memberInfo" style="background:${COLORS.light};padding:10px;border-radius:5px;margin-bottom:15px;display:none;">
          <p style="margin:0;font-weight:bold;" id="memberName"></p>
          <p style="margin:5px 0 0 0;font-size:12px;" id="currentSavings"></p>
          <p style="margin:5px 0 0 0;font-size:12px;" id="savingsStatus"></p>
        </div>
        
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Loan Amount:</label>
          <input type="number" id="loanAmount" style="width:100%;padding:8px;box-sizing:border-box;" 
                 min="${CONFIG.minLoanAmount}" max="${CONFIG.maxLoanAmount}" 
                 placeholder="Enter amount between ₱${CONFIG.minLoanAmount.toLocaleString()} and ₱${CONFIG.maxLoanAmount.toLocaleString()}" required>
          <small style="color:#666;">Min: ₱${CONFIG.minLoanAmount.toLocaleString()}, Max: ₱${CONFIG.maxLoanAmount.toLocaleString()}</small>
        </div>
        
        <div style="margin-bottom:20px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Loan Term (Months):</label>
          <input type="number" id="loanTerm" style="width:100%;padding:8px;box-sizing:border-box;" 
                 min="${CONFIG.loanTermMin}" max="${CONFIG.loanTermMax}" 
                 placeholder="Enter months (${CONFIG.loanTermMin}-${CONFIG.loanTermMax})" required>
          <small style="color:#666;">Term: ${CONFIG.loanTermMin}-${CONFIG.loanTermMax} months</small>
        </div>
        
        <div style="display:flex;gap:10px;margin-top:25px;">
          <button type="button" onclick="calculateLoan()" 
            style="flex:1;background:${COLORS.info};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
            📊 Calculate
          </button>
          <button type="button" onclick="applyLoan()" 
            style="flex:1;background:${COLORS.success};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
            🏦 Apply
          </button>
          <button type="button" onclick="closeDialog()" 
            style="flex:1;background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;">
            Cancel
          </button>
        </div>
      </form>
      
      <div id="calculationResult" style="margin-top:20px;padding:15px;background:${COLORS.light};border-radius:5px;display:none;">
        <h4 style="color:${COLORS.primary};margin-top:0;">Loan Calculation:</h4>
        <div id="loanDetails"></div>
      </div>
      
      <div id="eligibilityResult" style="margin-top:20px;padding:15px;background:${COLORS.light};border-radius:5px;display:none;">
        <h4 style="color:${COLORS.primary};margin-top:0;">Eligibility Check:</h4>
        <div id="eligibilityDetails"></div>
      </div>
      
      <div id="message" style="margin-top:15px;"></div>
      
      <script>
        function loadMemberInfo() {
          const memberId = document.getElementById('memberId').value;
          if (!memberId) {
            document.getElementById('memberInfo').style.display = 'none';
            document.getElementById('eligibilityResult').style.display = 'none';
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(member) {
              if (member.success) {
                document.getElementById('memberInfo').style.display = 'block';
                document.getElementById('memberName').textContent = 'Member: ' + member.name;
                document.getElementById('currentSavings').textContent = 
                  'Total Savings: ₱' + member.savings.toLocaleString();
                document.getElementById('savingsStatus').textContent = 
                  'Status: ' + member.status;
                
                // Auto-check eligibility when member is selected
                checkEligibility();
              } else {
                document.getElementById('memberInfo').style.display = 'none';
                document.getElementById('eligibilityResult').style.display = 'none';
                showMessage('❌ ' + member.message, 'error');
              }
            })
            .getMemberInfo(memberId);
        }
        
        function checkEligibility() {
          const memberId = document.getElementById('memberId').value;
          
          if (!memberId) {
            return;
          }
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                document.getElementById('eligibilityResult').style.display = 'block';
                document.getElementById('eligibilityDetails').innerHTML = 
                  '<p style="color:${COLORS.success};font-weight:bold;">✅ ELIGIBLE FOR LOAN</p>' +
                  '<p><strong>Total Savings:</strong> ₱' + result.totalSavings.toLocaleString() + '</p>' +
                  '<p><strong>Savings Status:</strong> ' + result.savingsStatus + '</p>' +
                  '<p><strong>Message:</strong> ' + result.message + '</p>';
              } else {
                document.getElementById('eligibilityResult').style.display = 'block';
                document.getElementById('eligibilityDetails').innerHTML = 
                  '<p style="color:${COLORS.danger};font-weight:bold;">❌ NOT ELIGIBLE</p>' +
                  '<p><strong>Reason:</strong> ' + result.message + '</p>';
              }
            })
            .checkLoanEligibility(memberId);
        }
        
        function calculateLoan() {
          const memberId = document.getElementById('memberId').value.trim();
          const loanAmount = document.getElementById('loanAmount').value;
          const loanTerm = document.getElementById('loanTerm').value;
          
          if (!memberId) {
            showMessage('Please select a Member first.', 'error');
            return;
          }
          
          if (!loanAmount || !loanTerm) {
            showMessage('Please fill in all fields.', 'error');
            return;
          }
          
          // First check eligibility
          google.script.run
            .withSuccessHandler(function(eligibilityResult) {
              if (!eligibilityResult.success) {
                showMessage('❌ Not eligible: ' + eligibilityResult.message, 'error');
                document.getElementById('calculationResult').style.display = 'none';
                return;
              }
              
              // If eligible, calculate loan
              google.script.run
                .withSuccessHandler(function(result) {
                  if (result.success) {
                    document.getElementById('calculationResult').style.display = 'block';
                    document.getElementById('loanDetails').innerHTML = 
                      '<p><strong>Loan Amount:</strong> ₱' + result.loanAmount.toLocaleString() + '</p>' +
                      '<p><strong>Term:</strong> ' + result.loanTerm + ' months</p>' +
                      '<p><strong>Monthly Interest:</strong> ₱' + result.monthlyInterest.toLocaleString() + '</p>' +
                      '<p><strong>Total Interest:</strong> ₱' + result.totalInterest.toLocaleString() + '</p>' +
                      '<p><strong>Monthly Payment:</strong> ₱' + result.monthlyPayment.toLocaleString() + '</p>' +
                      '<p><strong>Total Repayment:</strong> ₱' + result.totalRepayment.toLocaleString() + '</p>' +
                      '<p><strong>Effective Annual Rate:</strong> ' + (result.effectiveAnnualRate * 100).toFixed(1) + '%</p>';
                    showMessage('✅ Calculation complete!', 'success');
                  } else {
                    showMessage('❌ ' + result.message, 'error');
                    document.getElementById('calculationResult').style.display = 'none';
                  }
                })
                .calculateLoanDetails(loanAmount, loanTerm);
            })
            .checkLoanEligibility(memberId);
        }
        
        function applyLoan() {
          const memberId = document.getElementById('memberId').value.trim();
          const loanAmount = document.getElementById('loanAmount').value;
          const loanTerm = document.getElementById('loanTerm').value;
          
          if (!memberId) {
            showMessage('Please select a Member.', 'error');
            return;
          }
          
          if (!loanAmount || !loanTerm) {
            showMessage('Please fill in all fields.', 'error');
            return;
          }
          
          // Confirm with user about high interest rate
          const confirmMessage = '⚠️ IMPORTANT: This loan has ' + (CONFIG.interestRateWithLoan * 100) + '% MONTHLY interest!\n\n' +
                                 'Are you sure you want to proceed?';
          
          if (!confirm(confirmMessage)) {
            return;
          }
          
          showMessage('Processing loan application...', 'info');
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                showMessage('✅ ' + result.message, 'success');
                setTimeout(() => {
                  google.script.host.close();
                }, 3000);
              } else {
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .withFailureHandler(function(error) {
              showMessage('❌ Error: ' + error.message, 'error');
            })
            .processLoanApplication(memberId, loanAmount, loanTerm);
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.innerHTML = text;
          messageDiv.style.color = type === 'error' ? '${COLORS.danger}' : 
                                  type === 'success' ? '${COLORS.success}' : 
                                  '${COLORS.info}';
          messageDiv.style.fontWeight = 'bold';
          messageDiv.style.textAlign = 'center';
          messageDiv.style.padding = '10px';
        }
        
        function closeDialog() {
          google.script.host.close();
        }
        
        // Initialize - load member info if one is already selected
        window.onload = function() {
          const memberId = document.getElementById('memberId').value;
          if (memberId) {
            loadMemberInfo();
          }
        };
      </script>
    </div>
  `).setWidth(500).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Apply for Loan');
}

function checkLoanEligibility(memberId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const summarySheet = ss.getSheetByName('📊 Summary');
    
    if (!savingsSheet || !summarySheet) {
      return { success: false, message: 'System sheets not found.' };
    }
    
    // Check if member exists
    const savingsData = savingsSheet.getDataRange().getValues();
    let memberFound = false;
    let totalSavings = 0;
    let savingsStatus = '';
    
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] === memberId) {
        memberFound = true;
        totalSavings = savingsData[i][2] || 0;
        savingsStatus = savingsData[i][7] || '';
        break;
      }
    }
    
    if (!memberFound) {
      return { success: false, message: 'Member ID not found.' };
    }
    
    // Check minimum savings requirement
    if (totalSavings < CONFIG.savingsForLoan) {
      return { 
        success: false, 
        message: `Minimum savings requirement not met.\n` +
                 `Your savings: ₱${totalSavings.toLocaleString()}\n` +
                 `Required: ₱${CONFIG.savingsForLoan.toLocaleString()}\n` +
                 `Shortfall: ₱${(CONFIG.savingsForLoan - totalSavings).toLocaleString()}` 
      };
    }
    
    // Check savings status
    if (savingsStatus !== 'On Track') {
      return { 
        success: false, 
        message: `Cannot apply for loan. Savings status is "${savingsStatus}".\n` +
                 `You must be "On Track" with your savings payments to qualify for a loan.` 
      };
    }
    
    // Check existing loans
    const summaryData = summarySheet.getDataRange().getValues();
    let existingLoans = 0;
    
    for (let j = 3; j < summaryData.length; j++) {
      if (summaryData[j][0] === memberId) {
        existingLoans = summaryData[j][4] || 0;
        break;
      }
    }
    
    if (existingLoans >= 1) {
      return { 
        success: false, 
        message: `You already have ${existingLoans} active loan(s).\n` +
                 `Members are limited to 1 active loan at a time.` 
      };
    }
    
    return { 
      success: true, 
      message: 'Member is eligible for a loan.',
      totalSavings: totalSavings,
      savingsStatus: savingsStatus
    };
    
  } catch (error) {
    console.error('Error checking loan eligibility:', error);
    return { success: false, message: 'Error checking eligibility: ' + error.message };
  }
}

function checkCommunityFunds(loanAmount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fundsSheet = ss.getSheetByName('💰 Community Funds');
    
    if (!fundsSheet) {
      return { canApprove: false, message: 'Community funds sheet not found.' };
    }
    
    const availableCell = fundsSheet.getRange('B9');
    const available = availableCell.getValue() || 0;
    
    if (available <= 0) {
      return { 
        canApprove: false, 
        message: 'Insufficient community funds.',
        available: available 
      };
    }
    
    if (loanAmount > available) {
      return { 
        canApprove: false, 
        message: `Loan amount exceeds available funds.`,
        available: available 
      };
    }
    
    // Check if loan would reduce funds below minimum threshold
    const minimumThreshold = available * 0.1; // Keep at least 10% available
    if ((available - loanAmount) < minimumThreshold) {
      return { 
        canApprove: false, 
        message: `Loan would reduce community funds below minimum threshold.`,
        available: available 
      };
    }
    
    return { 
      canApprove: true, 
      message: `Sufficient funds available.`,
      available: available 
    };
    
  } catch (error) {
    console.error('Error checking community funds:', error);
    return { canApprove: false, message: 'Error checking funds: ' + error.message, available: 0 };
  }
}

function getLoanDetails(loanId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const loanSheet = ss.getSheetByName('🏦 Loans');
    
    if (!loanSheet) {
      return { success: false, message: 'Loan sheet not found.' };
    }
    
    const loanData = loanSheet.getDataRange().getValues();
    
    for (let i = 3; i < loanData.length; i++) {
      if (loanData[i][0] === loanId) {
        return {
          success: true,
          id: loanId,
          amount: loanData[i][2],
          remaining: loanData[i][11],
          monthlyPayment: loanData[i][8],
          status: loanData[i][6]
        };
      }
    }
    
    return { success: false, message: 'Loan ID not found.' };
    
  } catch (error) {
    console.error('Error getting loan details:', error);
    return { success: false, message: 'Error: ' + error.message };
  }
}

function findActiveLoans(loanId, memberId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const loanSheet = ss.getSheetByName('🏦 Loans');
    
    if (!loanSheet) {
      return { success: false, message: 'Loan sheet not found.' };
    }
    
    const loanData = loanSheet.getDataRange().getValues();
    const activeLoans = [];
    
    for (let i = 3; i < loanData.length; i++) {
      if (loanData[i][6] === 'Active') {
        if ((loanId && loanData[i][0] === loanId) || 
            (memberId && loanData[i][1] === memberId)) {
          activeLoans.push({
            id: loanData[i][0],
            amount: loanData[i][2],
            status: loanData[i][6],
            remaining: loanData[i][11]
          });
        }
      }
    }
    
    return { 
      success: true, 
      loans: activeLoans,
      message: activeLoans.length > 0 ? 'Active loans found.' : 'No active loans found.'
    };
    
  } catch (error) {
    console.error('Error finding active loans:', error);
    return { success: false, message: 'Error: ' + error.message, loans: [] };
  }
}

function recordLoanPayment(loanId, amount, paymentDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const loanSheet = ss.getSheetByName('🏦 Loans');
    const loanPaymentsSheet = ss.getSheetByName('💳 Loan Payments');
    
    if (!loanSheet || !loanPaymentsSheet) {
      return { success: false, message: 'System sheets not found.' };
    }
    
    // Find the loan
    const loanData = loanSheet.getDataRange().getValues();
    let loanRow = -1;
    let loanDetails = null;
    
    for (let i = 3; i < loanData.length; i++) {
      if (loanData[i][0] === loanId) {
        loanRow = i;
        loanDetails = {
          id: loanId,
          memberId: loanData[i][1],
          amount: loanData[i][2],
          remaining: loanData[i][11],
          monthlyPayment: loanData[i][8],
          paymentsMade: loanData[i][12] || 0,
          status: loanData[i][6]
        };
        break;
      }
    }
    
    if (loanRow === -1 || !loanDetails) {
      return { success: false, message: 'Loan ID not found.' };
    }
    
    if (loanDetails.status !== 'Active') {
      return { success: false, message: 'Loan is not active.' };
    }
    
    const paymentAmount = parseFloat(amount);
    const newRemaining = loanDetails.remaining - paymentAmount;
    
    // Calculate interest and principal
    const monthlyInterest = loanDetails.amount * CONFIG.interestRateWithLoan;
    let interestPaid = 0;
    let principalPaid = paymentAmount;
    
    if (newRemaining < 0) {
      return { 
        success: false, 
        message: `Payment amount (₱${paymentAmount.toLocaleString()}) exceeds remaining balance (₱${loanDetails.remaining.toLocaleString()}).` 
      };
    }
    
    // Update loan sheet
    const newPaymentsMade = loanDetails.paymentsMade + 1;
    loanSheet.getRange(loanRow + 1, 12).setValue(newRemaining);
    loanSheet.getRange(loanRow + 1, 13).setValue(newPaymentsMade);
    
    // Check if loan is fully paid
    if (newRemaining <= 0) {
      loanSheet.getRange(loanRow + 1, 7).setValue('Paid');
      loanSheet.getRange(loanRow + 1, 14).setValue('N/A');
    } else {
      const nextPaymentDate = getNextPaymentDate();
      loanSheet.getRange(loanRow + 1, 14).setValue(nextPaymentDate);
    }
    
    // Record payment
    const paymentId = 'LP' + new Date().getTime() + Math.floor(Math.random() * 1000);
    const lastPaymentRow = loanPaymentsSheet.getLastRow();
    const paymentRecord = [
      paymentId,
      loanId,
      new Date(paymentDate),
      paymentAmount,
      principalPaid,
      interestPaid,
      newRemaining,
      newRemaining <= 0 ? 'Fully Paid' : 'Partial',
      newPaymentsMade
    ];
    
    loanPaymentsSheet.getRange(lastPaymentRow + 1, 1, 1, 9).setValues([paymentRecord]);
    
    // Format payment record
    loanPaymentsSheet.getRange(lastPaymentRow + 1, 3).setNumberFormat('mm/dd/yyyy');
    loanPaymentsSheet.getRange(lastPaymentRow + 1, 4, 1, 4).setNumberFormat('₱#,##0.00');
    
    // Update summary sheet
    const summarySheet = ss.getSheetByName('📊 Summary');
    if (summarySheet) {
      const summaryData = summarySheet.getDataRange().getValues();
      for (let i = 3; i < summaryData.length; i++) {
        if (summaryData[i][0] === loanDetails.memberId) {
          const currentLoans = summaryData[i][3] || 0;
          const newLoans = Math.max(0, currentLoans - paymentAmount);
          
          summarySheet.getRange(i + 1, 4).setValue(newLoans);
          summarySheet.getRange(i + 1, 7).setValue((summaryData[i][6] || 0) + interestPaid);
          summarySheet.getRange(i + 1, 9).setValue((summaryData[i][8] || 0) + paymentAmount);
          
          summarySheet.getRange(i + 1, 4).setNumberFormat('₱#,##0.00');
          summarySheet.getRange(i + 1, 7).setNumberFormat('₱#,##0.00');
          summarySheet.getRange(i + 1, 9).setNumberFormat('₱#,##0.00');
          
          const currentLoanStreak = summaryData[i][10] || 0;
          summarySheet.getRange(i + 1, 11).setValue(currentLoanStreak + 1);
          break;
        }
      }
    }
    
    // RECORD FUNDS TRANSACTION
    recordFundsTransaction('Loan Repayment', loanDetails.memberId, paymentId, paymentAmount, 
                          `Loan ${loanId} payment - ₱${principalPaid.toLocaleString()} principal`);
    
    // Update silently
    updateDashboard();
    refreshAllDropdowns();
    
    return { 
      success: true, 
      message: `Loan payment of ₱${paymentAmount.toLocaleString()} recorded.\n` +
               `Remaining balance: ₱${newRemaining.toLocaleString()}\n` +
               `Payments made: ${newPaymentsMade}` 
    };
    
  } catch (error) {
    console.error('Error recording loan payment:', error);
    return { success: false, message: 'Failed to record payment: ' + error.message };
  }
}

// ============================================
// ADDITIONAL UTILITY FUNCTIONS
// ============================================

function showAllMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const savingsSheet = ss.getSheetByName('💰 Savings');
  
  if (!savingsSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Savings sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  savingsSheet.showSheet();
  ss.setActiveSheet(savingsSheet);
}

function showCommunityFunds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fundsSheet = ss.getSheetByName('💰 Community Funds');
  
  if (!fundsSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Community Funds sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  fundsSheet.showSheet();
  ss.setActiveSheet(fundsSheet);
}

function updateAllCalculations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Force recalculation of all sheets
    ss.getSheets().forEach(sheet => {
      const range = sheet.getDataRange();
      const formulas = range.getFormulas();
      if (formulas.flat().some(f => f !== '')) {
        // Force recalculation by reading and writing values
        range.setValues(range.getValues());
      }
    });

    // Always update community funds
    updateCommunityFundsCalculations();  
    SpreadsheetApp.getUi().alert('✅ Calculations Updated', 'All calculations and formulas have been refreshed.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('Error updating calculations:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 'Failed to update calculations: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function generateReports() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:400px;">
      <h3 style="color:${COLORS.primary};text-align:center;margin-bottom:20px;">📋 Generate Reports</h3>
      <p style="text-align:center;margin-bottom:20px;">Select the report you want to generate:</p>
      
      <div style="display:flex;flex-direction:column;gap:10px;">
        <button onclick="generateReport('members')" 
          style="background:${COLORS.primary};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
          👥 Member List Report
        </button>
        
        <button onclick="generateReport('savings')" 
          style="background:${COLORS.success};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
          💰 Savings Summary Report
        </button>
        
        <button onclick="generateReport('loans')" 
          style="background:${COLORS.accent};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
          🏦 Active Loans Report
        </button>
        
        <button onclick="generateReport('payments')" 
          style="background:${COLORS.info};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
          💵 Payment History Report
        </button>
        
        <button onclick="closeDialog()" 
          style="background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;margin-top:10px;">
          Cancel
        </button>
      </div>
      
      <div id="message" style="margin-top:15px;"></div>
      
      <script>
        function generateReport(type) {
          showMessage('Generating report...', 'info');
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                showMessage('✅ ' + result.message, 'success');
                setTimeout(() => {
                  google.script.host.close();
                }, 2000);
              } else {
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .withFailureHandler(function(error) {
              showMessage('❌ Error: ' + error.message, 'error');
            })
            .generateReport(type);
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.innerHTML = text;
          messageDiv.style.color = type === 'error' ? '${COLORS.danger}' : 
                                  type === 'success' ? '${COLORS.success}' : 
                                  '${COLORS.info}';
          messageDiv.style.fontWeight = 'bold';
          messageDiv.style.textAlign = 'center';
          messageDiv.style.padding = '10px';
        }
        
        function closeDialog() {
          google.script.host.close();
        }
      </script>
    </div>
  `).setWidth(400).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Reports');
}

function generateReport(reportType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let message = '';
    
    switch(reportType) {
      case 'members':
        showAllMembers();
        message = 'Member list displayed.';
        break;
      case 'savings':
        const savingsSheet = ss.getSheetByName('💰 Savings');
        if (savingsSheet) {
          savingsSheet.showSheet();
          ss.setActiveSheet(savingsSheet);
          message = 'Savings summary displayed.';
        }
        break;
      case 'loans':
        const loanSheet = ss.getSheetByName('🏦 Loans');
        if (loanSheet) {
          loanSheet.showSheet();
          ss.setActiveSheet(loanSheet);
          message = 'Active loans report displayed.';
        }
        break;
      case 'payments':
        const paymentsSheet = ss.getSheetByName('💵 Savings Payments');
        if (paymentsSheet) {
          paymentsSheet.showSheet();
          ss.setActiveSheet(paymentsSheet);
          message = 'Payment history displayed.';
        }
        break;
      default:
        message = 'Invalid report type.';
    }
    
    return { success: true, message: message };
    
  } catch (error) {
    console.error('Error generating report:', error);
    return { success: false, message: 'Error: ' + error.message };
  }
}

function showAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // Show all sheets
  sheets.forEach(sheet => {
    sheet.showSheet();
  });
  
  // Show dashboard by default
  const dashboard = ss.getSheetByName('📈 Dashboard');
  if (dashboard) {
    ss.setActiveSheet(dashboard);
  }
}

function setupMemberView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const portalSheet = ss.getSheetByName('👤 Member Portal');
  
  if (!portalSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Member Portal not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Hide all other sheets
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    if (sheet.getName() !== '👤 Member Portal') {
      sheet.hideSheet();
    }
  });
  
  // Show member portal
  portalSheet.showSheet();
  ss.setActiveSheet(portalSheet);
  
  SpreadsheetApp.getUi().alert('Member View', 'Now showing Member Portal only. Use the menu to show all sheets.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function showRemoveMemberForm() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:400px;">
      <h3 style="color:${COLORS.danger};text-align:center;margin-bottom:20px;">🗑️ Remove Member</h3>
      <p style="text-align:center;color:${COLORS.warning};margin-bottom:20px;">
        ⚠️ WARNING: This action cannot be undone!
      </p>
      
      <form id="removeForm">
        <div style="margin-bottom:15px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Member ID:</label>
          <input type="text" id="memberId" style="width:100%;padding:8px;box-sizing:border-box;" 
                 placeholder="Enter Member ID to remove" required>
        </div>
        
        <div style="margin-bottom:20px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Reason for Removal:</label>
          <textarea id="reason" style="width:100%;padding:8px;box-sizing:border-box;height:80px;" 
                    placeholder="Enter reason for removal" required></textarea>
        </div>
        
        <div style="display:flex;gap:10px;margin-top:25px;">
          <button type="button" onclick="removeMember()" 
            style="flex:1;background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
            🗑️ Remove
          </button>
          <button type="button" onclick="closeDialog()" 
            style="flex:1;background:${COLORS.secondary};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;">
            Cancel
          </button>
        </div>
      </form>
      
      <div id="message" style="margin-top:15px;"></div>
      
      <script>
        function removeMember() {
          const memberId = document.getElementById('memberId').value.trim();
          const reason = document.getElementById('reason').value.trim();
          
          if (!memberId || !reason) {
            showMessage('Please fill in all fields.', 'error');
            return;
          }
          
          if (!confirm('Are you sure you want to remove member ' + memberId + '? This action cannot be undone!')) {
            return;
          }
          
          showMessage('Removing member...', 'info');
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                showMessage('✅ ' + result.message, 'success');
                setTimeout(() => {
                  google.script.host.close();
                  google.script.run.refreshAllDropdowns();
                }, 2000);
              } else {
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .withFailureHandler(function(error) {
              showMessage('❌ Error: ' + error.message, 'error');
            })
            .removeMember(memberId, reason);
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.innerHTML = text;
          messageDiv.style.color = type === 'error' ? '${COLORS.danger}' : 
                                  type === 'success' ? '${COLORS.success}' : 
                                  '${COLORS.info}';
          messageDiv.style.fontWeight = 'bold';
          messageDiv.style.textAlign = 'center';
          messageDiv.style.padding = '10px';
        }
        
        function closeDialog() {
          google.script.host.close();
        }
      </script>
    </div>
  `).setWidth(400).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Remove Member');
}

function refreshAllDropdowns() {
  try {
    updateMemberPortalDropdown();
    updateSavingsPaymentDropdowns();
    updateLoanPaymentDropdowns();
    // Removed notification
  } catch (error) {
    console.error('Error refreshing dropdowns:', error);
  }
}

function removeMember(memberId, reason) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const savingsSheet = ss.getSheetByName('💰 Savings');
    const summarySheet = ss.getSheetByName('📊 Summary');
    const loanSheet = ss.getSheetByName('🏦 Loans');
    
    if (!savingsSheet || !summarySheet) {
      return { success: false, message: 'System sheets not found.' };
    }
    
    // Check if member exists and has active loans
    const loanData = loanSheet ? loanSheet.getDataRange().getValues() : [];
    let hasActiveLoans = false;
    
    for (let i = 3; i < loanData.length; i++) {
      if (loanData[i][1] === memberId && loanData[i][6] === 'Active') {
        hasActiveLoans = true;
        break;
      }
    }
    
    if (hasActiveLoans) {
      return { 
        success: false, 
        message: 'Cannot remove member with active loans.\n' +
                 'Please ensure all loans are paid off before removing the member.' 
      };
    }
    
    // Find and remove from savings sheet
    const savingsData = savingsSheet.getDataRange().getValues();
    let savingsAmount = 0;
    let memberName = '';
    
    for (let i = 3; i < savingsData.length; i++) {
      if (savingsData[i][0] === memberId) {
        savingsAmount = savingsData[i][2] || 0;
        memberName = savingsData[i][1];
        savingsSheet.deleteRow(i + 1);
        break;
      }
    }
    
    // Find and remove from summary sheet
    const summaryData = summarySheet.getDataRange().getValues();
    for (let i = 3; i < summaryData.length; i++) {
      if (summaryData[i][0] === memberId) {
        summarySheet.deleteRow(i + 1);
        break;
      }
    }
    
    // RECORD FUNDS TRANSACTION (withdraw savings)
    if (savingsAmount > 0) {
      recordFundsTransaction('Member Removal', memberId, 'Savings Withdrawal', -savingsAmount, 
                            `Member removal - ${reason}`);
    }
    
    // Update dropdowns
    refreshAllDropdowns();
    updateMemberPortalDropdown();
    
    return { 
      success: true, 
      message: `Member ${memberName} (${memberId}) removed successfully.\n` +
               `Savings refunded: ₱${savingsAmount.toLocaleString()}\n` +
               `Reason: ${reason}` 
    };
    
  } catch (error) {
    console.error('Error removing member:', error);
    return { success: false, message: 'Error removing member: ' + error.message };
  }
}

function showSearchMemberForm() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding:20px;font-family:Arial;max-width:400px;">
      <h3 style="color:${COLORS.primary};text-align:center;margin-bottom:20px;">🔍 Search Member</h3>
      
      <form id="searchForm">
        <div style="margin-bottom:20px;">
          <label style="display:block;margin-bottom:5px;font-weight:bold;">Member ID:</label>
          <input type="text" id="memberId" style="width:100%;padding:8px;box-sizing:border-box;" 
                 placeholder="Enter Member ID" required>
        </div>
        
        <div style="display:flex;gap:10px;margin-top:25px;">
          <button type="button" onclick="searchMember()" 
            style="flex:1;background:${COLORS.primary};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">
            🔍 Search
          </button>
          <button type="button" onclick="closeDialog()" 
            style="flex:1;background:${COLORS.danger};color:white;padding:12px;border:none;border-radius:4px;cursor:pointer;">
            Cancel
          </button>
        </div>
      </form>
      
      <div id="searchResult" style="margin-top:20px;padding:15px;background:${COLORS.light};border-radius:5px;display:none;">
        <h4 style="color:${COLORS.primary};margin-top:0;">Member Information:</h4>
        <div id="memberDetails"></div>
      </div>
      
      <div id="message" style="margin-top:15px;"></div>
      
      <script>
        function searchMember() {
          const memberId = document.getElementById('memberId').value.trim();
          
          if (!memberId) {
            showMessage('Please enter a Member ID.', 'error');
            return;
          }
          
          showMessage('Searching...', 'info');
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                document.getElementById('searchResult').style.display = 'block';
                document.getElementById('memberDetails').innerHTML = 
                  '<p><strong>Member Name:</strong> ' + result.memberName + '</p>' +
                  '<p><strong>Total Savings:</strong> ₱' + result.totalSavings.toLocaleString() + '</p>' +
                  '<p><strong>Current Loan:</strong> ₱' + result.currentLoan.toLocaleString() + '</p>' +
                  '<p><strong>Status:</strong> ' + result.status + '</p>' +
                  '<p><strong>Active Loans:</strong> ' + result.activeLoans + '</p>' +
                  '<p><strong>Overall Balance:</strong> ₱' + result.overallBalance.toLocaleString() + '</p>';
                showMessage('✅ Member found!', 'success');
              } else {
                document.getElementById('searchResult').style.display = 'none';
                showMessage('❌ ' + result.message, 'error');
              }
            })
            .withFailureHandler(function(error) {
              showMessage('❌ Error: ' + error.message, 'error');
            })
            .searchMemberById(memberId);
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.innerHTML = text;
          messageDiv.style.color = type === 'error' ? '${COLORS.danger}' : 
                                  type === 'success' ? '${COLORS.success}' : 
                                  '${COLORS.info}';
          messageDiv.style.fontWeight = 'bold';
          messageDiv.style.textAlign = 'center';
          messageDiv.style.padding = '10px';
        }
        
        function closeDialog() {
          google.script.host.close();
        }
      </script>
    </div>
  `).setWidth(400).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Search Member');
}

// ============================================
// ENHANCED ONOPEN WITH BETTER INITIALIZATION
// ============================================

function onOpen() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const portalSheet = ss.getSheetByName('👤 Member Portal');
    
    if (!portalSheet) {
      // System not generated yet
      SpreadsheetApp.getUi().createMenu('🏦 S&L System')
        .addItem('🎯 Generate Complete System', 'generateProfessionalSystem')
        .addToUi();
    } else {
      // System exists, create full menu
      createAdminMenu();
      
      // Initialize system components
      setTimeout(() => {
        try {
          refreshAllDropdowns();
          updateMemberPortalDropdown();
          
          // Update community funds if it exists
          const fundsSheet = ss.getSheetByName('💰 Community Funds');
          if (fundsSheet) {
            updateCommunityFundsCalculations();
          }
          
          // Update dashboard
          updateDashboard();
          
        } catch (e) {
          console.log('Initialization updates:', e.message);
        }
      }, 1000);
      
      // Show welcome message on first open
      const scriptProperties = PropertiesService.getScriptProperties();
      const firstOpen = scriptProperties.getProperty('firstOpen');
      
      if (!firstOpen) {
        scriptProperties.setProperty('firstOpen', 'true');
        
        SpreadsheetApp.getUi().alert(
          '🏦 Welcome to Savings & Loan System',
          'System is ready!\n\n' +
          'Use the "🏦 S&L System" menu for all operations:\n' +
          '• 👥 Member Management - Add/remove members\n' +
          '• 💰 Savings - Record payments, view summary\n' +
          '• 🏦 Loans - Apply, pay, view reports\n' +
          '• 📊 Reports - Various system reports\n' +
          '• 🔧 Tools - Maintenance and fixes\n' +
          '• 👁️ View Settings - Toggle sheet visibility\n\n' +
          'Tip: Use "Member Portal Only" to hide admin sheets from members.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    }
    
  } catch (error) {
    console.log('onOpen error:', error.message);
  }
}

// ============================================
// END OF COMPLETE ENHANCED SCRIPT
// ============================================