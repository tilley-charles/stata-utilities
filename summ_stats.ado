*! 1.0.0 CTilley Oct2017

capture program drop summ_stats
program              summ_stats
 qui {

  version 14.2
  syntax varlist using/ [if] [in], [sheet(string)] [replace]

  marksample touse, novarlist strok

  preserve

  keep if `touse'==1

  tempfile f
  postfile h str32(Variable) str80(Label) str1(Type) double(Mean Median SD Min Max nNonMiss nMiss pctMiss N) using "`f'"

  foreach var of varlist `varlist' {

    count if mi(`var')
    local nMiss = r(N)
    local nNonMiss = `c(N)' - r(N)

    local pctMiss = `nMiss'/`c(N)'*100

    if substr("`:type `var''",1,3) == "str" {
      post h ("`var'") ("`:variable label `var''") ("A") (.) (.) (.) (.) (.) (`nNonMiss') (`nMiss') (`pctMiss') (`c(N)')
    }
    else {
      summ `var', detail
      post h ("`var'") ("`:variable label `var''") ("N") (r(mean)) (r(p50)) (r(sd)) (r(min)) (r(max)) (`nNonMiss') (`nMiss') (`pctMiss') (`c(N)')
    }
   
  }

  postclose h
  use "`f'", clear

  tokenize `c(ALPHA)'

  export excel using "`using'", sheet("`sheet'") firstrow(variables) `=cond("`replace'"~="", "replace", "sheetreplace")'
  xlcolwidth   using "`using'", sheet("`sheet'") left(1 2 3) right(1 2 `c(k)') width(20 65 9)
  xlfreeze     using "`using'", sheet("`sheet'") row(1)
  putexcel set       "`using'", sheet("`sheet'") modify
  putexcel A1:``c(k)''1, bold
  putexcel A1:``c(k)''`=`c(N)'+1', font(Calibri, 10)
  putexcel C1:``c(k)''`=`c(N)'+1', hcenter
  putexcel D1:H`=`c(N)'+1', nformat(#,##0.0)
  putexcel I1:J`=`c(N)'+1', nformat(#,###)
  putexcel K1:K`=`c(N)'+1', nformat(#,##0.0)
  putexcel ``c(k)''1:``c(k)''`=`c(N)'+1', nformat(#,###)

  restore

 }

end

