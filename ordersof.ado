*! 1.0.0 CTilley 13Sep2019

capture program drop ordersof
program              ordersof
 qui {
  version 15.1

  syntax varname [if] [in], [clean] [local(name local)] [MISSing]

  if mi("`missing'") {
    marksample touse, strok
  }
  else {
    marksample touse, strok novarlist
  }

  capture confirm numeric variable `varlist'
  local numeric = (_rc==0)

  local clean = (~mi("`clean'"))

  tempvar tag
  egen   `tag' = tag(`varlist') if `touse'==1, `missing'

  count if `touse'==1

  if `numeric'==1 {
    local isint = (`varlist'==floor(`varlist'))
    mata: ordersofReal("`varlist'", "`tag'", `r(N)', `isint')
  }
  else {
    mata: ordersofString("`varlist'", "`tag'", `r(N)', `clean')
  }

  noi di as text `"`r(orders)'"'

  if ~mi("`local'") {
    c_local `local' = `"`r(orders)'"'
  }

 }
end

version 15.1
mata:
mata set matastrict on

void ordersofString(string scalar varname,
                    string scalar touse,
                    real   scalar N,
                    real   scalar clean)
{
  string colvector orders
  string scalar    sep
  string scalar    result

  st_sview(orders, ., varname, touse)

  if (clean==0) {
    sep = `"`=char(34)'"' + "`=char(39)'" + " " + "`=char(96)'" + `"`=char(34)'"'
  }
  else {
    sep = " "
  }

  result = invtokens(orders', sep)

  if (clean==0) {
    result = "`=char(96)'" + `"`=char(34)'"' + result + `"`=char(34)'"' + "`=char(39)'"
  }

  st_rclear()
  st_numscalar("r(N)", N)
  st_numscalar("r(r)", length(orders))
  st_global("r(orders)", result)
}

void ordersofReal(string scalar varname,
                  string scalar touse,
                  real   scalar N,
                  real   scalar isint)
{
  real   colvector orders
  string scalar    sep
  string scalar    tmp
  real   scalar    iter
  string scalar    result

  st_view(orders, ., varname, touse)

  sep = " "

  st_rclear()
  st_numscalar("r(N)", N)
  st_numscalar("r(r)", length(orders))

  if (isint==1) {
    result = invtokens(strofreal(orders', "%21.0g"), sep)
    st_global("r(orders)", result)
  }

  else {
    st_local("sep", sep)
    tmp = st_tempname()
    stata("local orders")
    for (iter = 1; iter <= rows(orders); iter++) {
      st_numscalar(tmp, orders[iter])
      stata("local el = " + tmp)
      stata("local orders" + "`=char(96)'" + "sep" + "`=char(39)'" + "`=char(96)'" + "orders" + "`=char(39)'" + "`=char(96)'" + "sep" + "`=char(39)'" + "`=char(96)'" + "el" + "`=char(39)'")
    }
    st_global("r(orders)", st_local("orders"))
  }
}

end
