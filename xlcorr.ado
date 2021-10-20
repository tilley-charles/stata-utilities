*! 1.0.0 CTilley Dec2019

capture program drop xlcorr
program define       xlcorr
 qui {
  version 16.0
  
  syntax varlist(min = 2 numeric) using/, [replace sheet(string) width(integer 15) nformat(string)]

  preserve

    * pairwise correlation
    capture pwcorr `varlist'
    if _rc~=0 {
      exit _rc
    }

    clear

    tempname rho
    matrix  `rho' = r(C)
    svmat   `rho', names(col)

    * export correlation matrix
    export excel using "`using'", `replace' sheet(`sheet') cell(B1) firstrow(variables)

    * formatting setup
    local k = c(k) + 1
    mata: st_local("K", numtobase26(`k'))
    local N = c(N) + 1

    ds
    local colnames `r(varlist)'

    * formatting
    putexcel set "`using'", `sheet' modify

    local nformat = cond(~mi("`nformat'"), "`nformat'", "##0.000")

    putexcel B2:`K'`N', hcenter nformat("`nformat'")

    * highlighting
    forvalues c = 1/`c(k)' {

      local v : word `c' of `colnames'
      mata: st_local("X", numtobase26(`=`c'+1'))

      forvalues r = 1/`c(N)' {

        local Y = `r' + 1

        if      `v'[`r']>= 0.0 & `v'[`r']<0.1             putexcel `X'`Y', fpattern(solid, "229 245 224")
        else if `v'[`r']>= 0.1 & `v'[`r']<0.2             putexcel `X'`Y', fpattern(solid, "199 233 192")
        else if `v'[`r']>= 0.2 & `v'[`r']<0.3             putexcel `X'`Y', fpattern(solid, "161 217 155")
        else if `v'[`r']>= 0.3 & `v'[`r']<0.4             putexcel `X'`Y', fpattern(solid, "116 196 118")
        else if `v'[`r']>= 0.4 & `v'[`r']<(1-c(epsfloat)) putexcel `X'`Y', fpattern(solid, " 65 171  93")
        else if `v'[`r']<  0.0 & `v'[`r']>-0.1            putexcel `X'`Y', fpattern(solid, "254 224 210")
        else if `v'[`r']<=-0.1 & `v'[`r']>-0.2            putexcel `X'`Y', fpattern(solid, "252 187 161")
        else if `v'[`r']<=-0.2 & `v'[`r']>-0.3            putexcel `X'`Y', fpattern(solid, "252 146 114")
        else if `v'[`r']<=-0.3 & `v'[`r']>-0.4            putexcel `X'`Y', fpattern(solid, "251 106  74")
        else if `v'[`r']<=-0.4                            putexcel `X'`Y', fpattern(solid, "239  59  44")
    
      }

    }    

    putexcel close

    * export row names
    describe, clear replace

    export excel name using "`using'", sheet(`sheet', modify) cell(A2)

    * column widths
    capture which xlcolwidth
    if _rc==0 {
      xlcolwidth using "`using'", `sheet' left(1) right(`k') width(`width')
    }

    * freeze panes
    capture which xlfreeze
    if _rc==0 {
      xlfreeze using "`using'", `sheet' row(1) col(1)
    }

  restore

 }  
end

