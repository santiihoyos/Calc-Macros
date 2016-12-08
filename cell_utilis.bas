REM  *****  BASIC  *****

Sub Main

End Sub

REM  *****  BASIC  *****
REM ################### RETURNING STRING #################################################
Function CELL_NOTE(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
REM returns annotation text
Dim v
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      CELL_NOTE = v.Annotation.getText.getString
   else
      CELL_NOTE = v
   endif
End Function

Function CELL_URL(vSheet,lRowIndex&,iColIndex%,optional n%)
'calls: getSheetCell
REM returns URL of Nth text-hyperlink from a cell, default N=1)
Dim v
   If isMissing(n) then n= 1
   If n < 1 then
      CELL_URL = Null
      exit function
   endif
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      if v.Textfields.Count >= n  then 
         CELL_URL = v.getTextfields.getByIndex(n -1).URL 
      else
         Cell_URL = Null
      endif
   else
      CELL_URL = v
   endif
End Function

Function CELL_FORMULA(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
REM return unlocalized (English) formula
Dim v
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      CELL_FORMULA = v.getFormula()
   else
      CELL_FORMULA = v
   endif
End Function

Function CELL_STYLE(vSheet,lRowIndex&,iColIndex%,optional bLocalized)
'calls: getSheetCell
REM return name of cell-style, optionally localized
Dim v,s$,bLocal as Boolean
   if not isMissing(bLocalized) then bLocal=cBool(bLocalized)
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      if bLocal then
         s = thisComponent.StyleFamilies("CellStyles").getByName(v.CellStyle).DisplayName
      else
         s = v.CellStyle
      endif
      CELL_STYLE = s
   else
      CELL_STYLE = v
   endif
End Function

Function CELL_LINE(vSheet,lRowIndex&,iColIndex%,optional n)
'calls: getSheetCell
REM Split by line breaks, missing or zero line number returns whole string.
REM =CELL_LINE(SHEET(),1,1,2) -> second line of A1 in this sheet
Dim v,s$,a(),i%
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      s = v.getString
      if not isMissing(n) then i = cInt(n)
      if i > 0 then
         a() = Split(s,chr(10))
         If (i <= uBound(a())+1)then
            CELL_LINE = a(i -1)
         else
            CELL_LINE = NULL
         endif
      else
         CELL_LINE = s
      endif
   else
      CELL_LINE = v
   endif
end Function

REM ################### RETURNING NUMBER #################################################
Function CELL_ISHORIZONTALPAGEBREAK(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
Dim v
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      CELL_ISHORIZONTALPAGEBREAK = Abs(cINT(v.Rows.getByIndex(0).IsStartOfNewPage))
   else
      CELL_ISHORIZONTALPAGEBREAK = v
   endif
End Function

Function CELL_ISVERTICALPAGEBREAK(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
Dim v
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      CELL_ISVERTICALPAGEBREAK = Abs(cINT(v.Columns.getByIndex(0).IsStartOfNewPage))
   else
      CELL_ISVERTICALPAGEBREAK = v
   endif
End Function

Function CELL_CHARCOLOR(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
REM returns color code as number
Dim v
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      CELL_CHARCOLOR = v.CharColor
   else
      CELL_CHARCOLOR = v
   endif
End Function
Function CELL_BACKCOLOR(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
REM returns color code as number
Dim v
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      CELL_BACKCOLOR = v.CellBackColor
   else
      CELL_BACKCOLOR = v
   endif
End Function
Function CELL_VISIBLE(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
REM returns visibility state as number 0|1
Dim v
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      CELL_VISIBLE = Abs(v.Rows.isVisible)
   else
      CELL_VISIBLE = v
   endif
End Function
Function CELL_LOCKED(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
REM returns locked state as number 0|1
Dim v
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      CELL_LOCKED = Abs(v.CellProtection.isLocked)
   else
      CELL_LOCKED = v
   endif
End Function
Function CELL_NumberFormat(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
REM returns the number format index
Dim v
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      CELL_NumberFormat = v.NumberFormat
   else
      CELL_NumberFormat = v
   endif
End Function
Function CELL_NumberFormatType(vSheet,lRowIndex&,iColIndex%)
'calls: getSheetCell
REM return a numeric com.sun.star.util.NumberFormat which describes a format category
Dim v,lNF&
   v = getSheetCell(vSheet,lRowIndex&,iColIndex%)
   if vartype(v) = 9 then
      lNF = v.NumberFormat
      CELL_NumberFormatType = ThisComponent.getNumberFormats.getByKey(lNF).Type
   else
      CELL_NumberFormatType = v
   endif
End Function

'################### HELPERS FOR ABOVE CELL FUNCTIONS #########################################
Function getSheet(byVal vSheet)
REM Helper for sheet functions. Get cell from sheet's name or position; cell's row-position; cell's col-position
on error goto exitErr
   select case varType(vSheet)
   case is = 8
      if thisComponent.sheets.hasbyName(vSheet) then
         getSheet = thisComponent.sheets.getByName(vSheet)
      else
         getSheet = NULL
      endif
   case 2 to 5
      vSheet = cInt(vSheet)
      'Wow! Calc has sheets with no name at index < 0,
      ' so NOT isNull(oSheet), if vSheet <= lbound(sheets) = CRASH!
      'http://www.openoffice.org/issues/show_bug.cgi?id=58796
      if(vSheet <= thisComponent.getSheets.getCount)AND(vSheet > 0) then
         getSheet = thisComponent.sheets.getByIndex(vSheet -1)
      else
         getSheet = NULL
      endif
   end select
exit function
exitErr:
getSheet = NULL
End Function

Function getSheetCell(byVal vSheet,byVal lRowIndex&,byVal iColIndex%)
dim oSheet
'   print vartype(vsheet)
   oSheet = getSheet(vSheet)
   if varType(oSheet) <>9 then
      getSheetCell = NULL
   elseif (lRowIndex > oSheet.rows.count)OR(lRowIndex < 1) then
      getSheetCell = NULL
   elseif (iColIndex > oSheet.columns.count)OR(iColIndex < 1) then
      getSheetCell = NULL
   else
      getSheetCell = oSheet.getCellByPosition(iColIndex -1,lRowIndex -1)
   endif
End Function

Sub Macro1

End Sub
