*! 1.2.0 CTilley 05Feb2018

capture program drop xlcolwidth
program xlcolwidth
  version 14.1

  syntax using/, [sheet(string)] Leftcol(numlist >0) Rightcol(numlist >0) Width(numlist >=0 <=255)


  *****************
  * Initial checks
  *****************

  local N : word count `leftcol'
	
  * leftcol, rightcol, and width must have same number of elements
  if (`N' ~= `: word count `rightcol'') | (`N' ~= `: word count `width'') {
    noi di as error "Options leftcol, rightcol, and width must contain the same number of elements"
    exit 198
  }
	
  * rightcol must be >= matching leftcol
  forvalues i = 1/`N' {
    local l : word `i' of `leftcol'
    local r : word `i' of `rightcol'
		
    if `r' < `l' {
      noi di as error "Right column must be >= left column"
      exit 198
    }
  }

  * xl sheet names cannot exceed 31 characters
  if strlen("`sheet'") > 31 {
    noi di as error "Excel sheet names cannot exceed 31 characters"
    exit 198
  }

  * xl sheet names cannot include certain characters
  if "`sheet'" ~= "" {
    foreach c in "\" "/" "?" "*" "[" "]" {
      if strpos("`sheet'", "`c'") ~= 0 {
        noi di as error "Sheet name includes illegal character: `c'"
        exit 198
      }
    }
  }
	
	
  *****************************
  * Loop through column widths
  *****************************
	
  forvalues i = 1/`N' {
    local l : word `i' of `leftcol'
    local r : word `i' of `rightcol'
    local w : word `i' of `width'
		
    if "`sheet'" ~= "" mata withsheet("`using'", "`sheet'", `l', `r', `w')
    else               mata nosheet(  "`using'",            `l', `r', `w')
  }
	
end

************************
* Define mata functions
************************

version 14.1
mata
mata clear

/* if sheet is specified */
void withsheet(xlfile, xlsheet, real scalar left, real scalar right, real scalar width) {
  class xl scalar b
	
  b = xl()
  b.set_mode("open")
  b.set_error_mode("off")
	
  b.load_book(xlfile)
  if (b.get_last_error() ~= 0) {
    b.create_book(xlfile, xlsheet, "xlsx")
  }

  b.set_sheet(xlsheet)
  if (b.get_last_error() ~= 0) {
    b.add_sheet(xlsheet)
    b.set_sheet(xlsheet)
  }
	
  b.set_column_width(left, right, width)
  b.close_book()
}

/* if sheet is not specified, will change widths on first sheet */
void nosheet(xlfile, real scalar left, real scalar right, real scalar width) {
  class xl scalar b
	
  b = xl()
  b.set_mode("open")
  b.set_error_mode("off")
	
  b.load_book(xlfile)
  if (b.get_last_error() ~= 0) {
    b.create_book(xlfile, "Sheet1", "xlsx")
  }
	
  b.set_column_width(left, right, width)
  b.close_book()
}

end
