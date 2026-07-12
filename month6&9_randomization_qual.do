
import excel "C:\Users\abubakar.farah_gived\Documents\Month6&9 qualitative work.xlsx", sheet("Sheet1") firstrow clear
tab ParentGeographicLevelName
replace ParentGeographicLevelName = "Kitengela" if inlist(ParentGeographicLevelName, "Kitengela_1", "Kitengela_2", "Kitengela_3", "Kitengela_4")
replace ParentGeographicLevelName = "Kasarani" if inlist(ParentGeographicLevelName, " Maji Mazuri", "Mikinduri", "Mugumoini")
replace ParentGeographicLevelName = "Kasarani" if inlist(ParentGeographicLevelName, "Mwiki")
replace ParentGeographicLevelName = "Kasarani" if inlist(ParentGeographicLevelName, "Maji Mazuri")
replace ParentGeographicLevelName = "Dagoretti" if inlist(ParentGeographicLevelName, "Dagoretti South", "Dagoretti North")
tab ParentGeographicLevelName
codebook,compact
tab IsRefugee
tab TreatmentGroup
tab CountryofOrigin
replace CountryofOrigin = "Kenya" if CountryofOrigin == "Kenya "
tab CountryofOrigin
import excel "C:\Users\abubakar.farah_gived\Documents\Month6&9 qualitative work.xlsx", sheet("Sheet1") firstrow clear
tab ParentGeographicLevelName
replace ParentGeographicLevelName = "Kitengela" if inlist(ParentGeographicLevelName, "Kitengela_1", "Kitengela_2", "Kitengela_3", "Kitengela_4")
replace ParentGeographicLevelName = "Kasarani" if inlist(ParentGeographicLevelName, " Maji Mazuri", "Mikinduri", "Mugumoini")
replace ParentGeographicLevelName = "Kasarani" if inlist(ParentGeographicLevelName, "Mwiki")
replace ParentGeographicLevelName = "Kasarani" if inlist(ParentGeographicLevelName, "Maji Mazuri")
replace ParentGeographicLevelName = "Dagoretti" if inlist(ParentGeographicLevelName, "Dagoretti South", "Dagoretti North")
tab ParentGeographicLevelName
replace CountryofOrigin = "Kenya" if CountryofOrigin == "Kenya "
codebook, compact
ds
tab AgeCalculated
gen youth = (AgeCalculated >= 18 & AgeCalculated <= 35) if !missing(AgeCalculated)
label define youth_lbl 0 "Non-Youth (36+)" 1 "Youth (18–35)"
label values youth youth_lbl
label variable youth "Youth (18–35)"
drop if AgeCalculated > 80
tab AgeCalculated
fre youth
list GDID LastName AgeCalculated if youth == 1
list GDID AgeCalculated if youth == 1
ParentGeographicLevelName AgeCalculated if youth == 1, replace
preserve
    set seed 20260429
    gen rand_temp = runiform() if youth == 1
    sort rand_temp
    keep if youth == 1
    keep in 1/100
    keep GDID IsRefugee Gender PrimaryContactNumber LastName  TreatmentGroup ParentGeographicLevelName CountryofOrigin AgeCalculated
    export excel using "youth_103.xlsx", firstrow(variables) replace
restore
tab youth
tab IsRefugee TreatmentGroup if youth == 1, m
tab ParentGeographicLevelName if youth == 1, m
tab CountryofOrigin if youth == 1, m
preserve
    set seed 20260430
    gen rand_temp2 = runiform() if youth == 0
    sort rand_temp2
    keep if youth == 0
    keep in 1/100
    keep GDID IsRefugee GenderLastName  TreatmentGroup ParentGeographicLevelName PrimaryContactNumber CountryofOrigin AgeCalculated
    export excel using "nonyouth_103.xlsx", firstrow(variables) replace
restore
