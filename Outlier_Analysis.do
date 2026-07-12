****************************************************/

clear all
set more off

import excel "C:\Users\abubakar.farah_gived\Downloads\Uganda FCDO Spending Analysis.xlsx", sheet("Analysis") firstrow

* Display basic information about the dataset
describe
summarize

/*******************************************************************************
* SECTION 2: VARIABLE RENAMING
*******************************************************************************/

* Rename long variable names to shorter ones
 rename Spending_Categories_Business_Oth SpendingBusiness
 rename Spending_Health SpendingHealth
******** mean and median analysis******
asdoc summ, detail stat(N mean sd min p50 max p1 p99)
asdoc summ, detail stat(mean sd p50)


* Verify all variables exist
local spending_vars "SpendingFood SpendingLivestock SpendingAgriculture SpendingBuildingHome SpendingHouseholdGoods SpendingOtherEducation SpendingTransport SpendingBusiness SpendingHealth"

foreach var of local spending_vars {
    capture confirm variable `var'
    if _rc {
        display as error "Warning: `var' not found in dataset"
    }
}

/*******************************************************************************
* SECTION 3: CREATE OUTLIER INDICATORS
*******************************************************************************/

display ""
display "Creating outlier indicators for all spending variables..."
display ""

foreach var of local spending_vars {
    
    * Check if variable exists
    capture confirm variable `var'
    if !_rc {
        
        * Drop existing outlier variables if they exist
        capture drop `var'_2sd
        capture drop `var'_3sd
        
        * Calculate statistics
        quietly summarize `var', detail
        local mean = r(mean)
        local sd = r(sd)
        local n = r(N)
        
        * Create outlier indicators
        gen `var'_2sd = (`var' > `mean' + 2*`sd') if `var' != .
        gen `var'_3sd = (`var' > `mean' + 3*`sd') if `var' != .
        
        * Label the variables
        label variable `var'_2sd "`var' outliers (>2 SD)"
        label variable `var'_3sd "`var' outliers (>3 SD)"
        
        * Display progress
        display "Created outlier variables for `var' (N=`n')"
    }
}

/*******************************************************************************
* SECTION 4: OUTLIER ANALYSIS
*******************************************************************************/

display ""
display "=========================================="
display "OUTLIER ANALYSIS RESULTS"
display "=========================================="
display ""

* Loop through each variable and display results
foreach var of local spending_vars {
    
    capture confirm variable `var'
    if !_rc {
        
        display ""
        display "--- `var' ---"
        
        * Basic statistics
        quietly summarize `var', detail
        display "N: " r(N)
        display "Mean: " %12.0fc r(mean)
        display "Median: " %12.0fc r(p50)
        display "SD: " %12.0fc r(sd)
        
        * Count outliers
        quietly sum `var'_2sd
        local out2 = r(sum)
        local pct2 = round(100*r(mean), 0.1)
        
        quietly sum `var'_3sd
        local out3 = r(sum)
        local pct3 = round(100*r(mean), 0.1)
        
        display "2SD outliers: `out2' (`pct2'%)"
        display "3SD outliers: `out3' (`pct3'%)"
        
        * Check if median > mean
        quietly summarize `var', detail
        if r(p50) > r(mean) {
            display "Note: Median > Mean (left-skewed)"
        }
        else {
            display "Note: Median < Mean (right-skewed)"
        }
    }
}

/*******************************************************************************
* SECTION 5: EXPORT TO EXCEL
*******************************************************************************/

display ""
display "Exporting results to Excel..."

* Create Excel file
putexcel set "Outlier_Analysis_Results.xlsx", replace

* Title and headers
putexcel A1 = "OUTLIER ANALYSIS REPORT - UGANDA CASH TRANSFER PROGRAM"
putexcel A1:I1, merge bold font(Arial, 14)

putexcel A2 = "Analysis Date: $S_DATE"
putexcel A2:I2, merge italic

putexcel A4 = "Variable"
putexcel B4 = "N (Obs)"
putexcel C4 = "Mean"
putexcel D4 = "Median"
putexcel E4 = "Std Dev"
putexcel F4 = "2SD Outliers"
putexcel G4 = "2SD %"
putexcel H4 = "3SD Outliers"
putexcel I4 = "3SD %"
putexcel A4:I4, bold border(bottom, thick)

* Data rows
local row = 5

foreach var of local spending_vars {
    
    capture confirm variable `var'
    if !_rc {
        
        * Calculate statistics
        quietly summarize `var', detail
        local n = r(N)
        local mean = r(mean)
        local median = r(p50)
        local sd = r(sd)
        
        * Count outliers
        quietly sum `var'_2sd
        local out2 = r(sum)
        local pct2 = round(100*r(mean), 0.1)
        
        quietly sum `var'_3sd
        local out3 = r(sum)
        local pct3 = round(100*r(mean), 0.1)
        
        * Write to Excel
        putexcel A`row' = "`var'"
        putexcel B`row' = `n'
        putexcel C`row' = `mean', nformat(#,##0)
        putexcel D`row' = `median', nformat(#,##0)
        putexcel E`row' = `sd', nformat(#,##0)
        putexcel F`row' = `out2'
        putexcel G`row' = `pct2', nformat(0.0)
        putexcel H`row' = `out3'
        putexcel I`row' = `pct3', nformat(0.0)
        
        local ++row
    }
}

* Format the table
putexcel A5:I`=`row'-1', border(all)

* Add notes
local note_row = `row' + 1
putexcel A`note_row' = "Notes:"
putexcel A`note_row', bold
local ++note_row
putexcel A`note_row' = "1. Outliers defined as values > mean + k*SD where k=2 or k=3"
local ++note_row
putexcel A`note_row' = "2. All values in Uganda Shillings (UGX)"
local ++note_row
putexcel A`note_row' = "3. Data has been winsorized at 3,700,000 UGX"
local ++note_row
putexcel A`note_row' = "4. Under normality, expect ~4.55% at 2SD and ~0.27% at 3SD"

/*******************************************************************************
* SECTION 6: CREATE SUMMARY TABLE IN STATA
*******************************************************************************/

* Create matrix for display
matrix outliers = J(9, 4, .)
matrix colnames outliers = "N" "2SD_Count" "2SD_%" "3SD_Count" "3SD_%"
matrix rownames outliers = "Food" "Livestock" "Agriculture" "Building" "Household" "Education" "Transport" "Business" "Health"

local i = 0
foreach var of local spending_vars {
    local ++i
    
    capture confirm variable `var'
    if !_rc {
        quietly sum `var'
        matrix outliers[`i', 1] = r(N)
        
        quietly sum `var'_2sd
        matrix outliers[`i', 2] = r(sum)
        matrix outliers[`i', 3] = round(100*r(mean), 0.1)
        
        quietly sum `var'_3sd
        matrix outliers[`i', 4] = r(sum)
        matrix outliers[`i', 5] = round(100*r(mean), 0.1)
    }
}

* Display matrix
display ""
display "Summary Table of Outliers:"
matlist outliers, format(%9.0f %9.0f %5.1f %9.0f %5.1f)




