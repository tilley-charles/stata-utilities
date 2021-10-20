*! 1.2.0 CTilley 05Oct2021

capture program drop xlfreeze
program define       xlfreeze
qui {
  version 15.1

  syntax using/, [Sheet(string)] [Row(integer 0)] [Column(integer 0)]

  local pshell %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe

  * store cd
  local cd `c(pwd)'
  cd "C:\"

  * file path normalization
  local using = subinstr("`using'", "/", "\", .)
  forvalues x = 1/10 {
    local using = subinstr("`using'", "\\", "\", .)
  }
  if substr("`using'", 1, 1)=="\" local using = "\" + "`using'"

  * make sure that file exists
  if strpos("`using'", "\")==0 local using "`cd'\\`using'"

  capture confirm file "`using'"
  if _rc~=0 {
    noi di as error "Excel workbook not found"
    exit 601
  }

  * make sure sheet exists
  if ~mi("`sheet'") {
    import excel using "`using'", describe
    local pass = 0
    forvalues w = 1/`r(N_worksheet)' {
      if strupper("`r(worksheet_`w')'")==strupper("`sheet'") {
        local pass = 1
        continue, break
      }
    }
    if `pass'==0 {
      noi di as error "Sheet {it:`sheet'} does not exist"
      exit 601
    }
  }

  * check that split rows and split columns are valid
  if (abs(trunc(`row'))==`row') + (abs(trunc(`column'))==`column') ~= 2 {
    noi di as error "Row and column arguments must be non-negative integers"
    exit 198
  }

  * write powershell script
  tempfile f
  local f = subinstr("`f'", "/", "\", .)
  local fname = strreverse(substr(strreverse("`f'"), 1, strpos(strreverse("`f'"), "\")-1))
  ! rename "`f'" `"`=subinstr("`fname'", ".tmp", ".ps1", .)'"'
  local f = subinstr("`f'", ".tmp", ".ps1", .)

  tempname   h
  file open `h' using "`f'", write text replace

    file write `h' `"\$filename = "`using'""' _n
  if ~mi("`sheet'") {
    file write `h' `"\$sheetname = "`sheet'""' _n
  }
    file write `h' `"\$excel = New-Object -comobject Excel.Application"' _n
    file write `h' `"\$wbook = \$excel.Workbooks.Open(\$filename)"' _n
  if ~mi("`sheet'") {
    file write `h' `"\$wsheet = \$wbook.Worksheets.Item(\$sheetname)"' _n
  }
  else {
    file write `h' `"\$wsheet = \$wbook.Worksheets.Item(1)"' _n 
  }
    file write `h' `"\$wsheet.Activate()"' _n
    file write `h' `"\$wsheet.Application.ActiveWindow.SplitRow = `row'"' _n
    file write `h' `"\$wsheet.Application.ActiveWindow.SplitColumn = `column'"' _n
  if ~(`row'==0 & `column'==0) {
    file write `h' `"\$wsheet.Application.ActiveWindow.FreezePanes = \$true"' _n  
  }
    file write `h' `"\$excel.DisplayAlerts = \$false"' _n
    file write `h' `"\$wbook.SaveAs(\$filename)"' _n
    file write `h' `"\$wbook.Close()"' _n    
    file write `h' `"\$excel.Quit"' _n

  file close `h'

  * execute script
  ! "`pshell'" "`f'"
  capt erase   "`f'"

  * return cd
  cd "`cd'"

}
end
