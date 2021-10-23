*! 1.0.2 CTilley 23Oct2021

capture program drop summ_stats
program              summ_stats
 qui {

  version 14.2
  syntax varlist using/ [if] [in], [sheet(string)] [replace]

  marksample touse, novarlist strok

  preserve

    keep if `touse'==1

    tempfile f
    postfile h str32(Variable) str80(Label) str1(Type) double(Mean P50 SD Min Max N nNonMiss nMiss pctMiss) using "`f'"

    /* summary statistics */
    foreach var of varlist `varlist' {

      count if mi(`var')
      local nMiss = r(N)
      local nNonMiss = `c(N)' - r(N)

      local pctMiss = `nMiss'/`c(N)'*100

      if substr("`:type `var''",1,3) == "str" {
        post h ("`var'") ("`:variable label `var''") ("A") `= 5 * `"(.) "'' (`c(N)') (`nNonMiss') (`nMiss') (`pctMiss')
      }
      else {
        summ `var', detail
        post h ("`var'") ("`:variable label `var''") ("N") (r(mean)) (r(p50)) (r(sd)) (r(min)) (r(max)) (`c(N)') (`nNonMiss') (`nMiss') (`pctMiss')
      }
     
    }

    /* output */
    postclose h
    use "`f'", clear

    tokenize `c(ALPHA)'

    export excel using "`using'", sheet("`sheet'") firstrow(variables) `=cond("`replace'"~="", "replace", "sheetreplace")'

    capture which xlcolwidth
    if _rc==0 {
      xlcolwidth using "`using'", sheet("`sheet'") left(1 2 3) right(1 2 `c(k)') width(20 40 9)
    }

    capture which xlfreeze
    if _rc==0 {
      xlfreeze using "`using'", sheet("`sheet'") row(1)
    }

    putexcel set       "`using'", sheet("`sheet'") modify
    putexcel A1:``c(k)''1, bold
    putexcel A1:``c(k)''`=`c(N)'+1', font(Calibri, 10)
    putexcel C1:``c(k)''`=`c(N)'+1', hcenter
    putexcel D1:H`=`c(N)'+1', nformat(#,##0.0)
    putexcel I1:K`=`c(N)'+1', nformat(#,##0)
    putexcel ``c(k)''1:``c(k)''`=`c(N)'+1', nformat(#,##0.0)

  restore

 }

end

