<%
  const constVBSdbVersion = "Pro-Revision.2003.12"

  const constRequestValueListDelimiter = "{_/{{__"  ' used as delimiter to build edit sql statement
  const constStrBlankReplacementsForInputFormControlNames = "__vdbla__"  ' if a field name containes blanks, than these are replaced with this constant, in input screens, to determine the related input form control name
  const constStrReplacementForQuotesInEditWhereConditions = "_VBSdb_DoubleQuotes_VBSdb_"  ' if a key field value contains quotes, these are replaced to avoid to mistakenly close the hidden EditWhere clause value
  'const constStrDotReplacementForInputFormControlNames = "__vddot__"  ' if a field name containes a dot it is replaced with this constant, in input screens, to determine the related input form control name
  const constDefaultMaxLenght = 25  ' default maxlength size for edit fields that are not text fields
  const constIntDefaultMemoItemEditModeNumRows = 5  ' default number of rows, for edit memo fields
  const constIntDefaultMemoItemEditModeNumCols = 42 ' default number of columns, for edit memo fields
  const constVbsDbGetOutputFieldValue_inWrongScreenType = "vbsDbConstVbsDbGetOutputFieldValue_inWrongScreenType"
  const constVbsDbGetOutputFieldValueEmptySqlRecordSet = "vbsDbConstVbsDbGetOutputFieldValueEmptySqlRecordSet"
  const constAutomaticSearchForAutoincrement = 0

  const constStrDefaultPrimarySeparator = ";"
  const constStrDefaultSecondarySeparator = "|"
  
  ' default color constants
  const constDefaultViewNavigationBGColor = "#d3d3d3"
  const constDefaultViewNavigationDisabledFGColor = "#9C9A9C"
  const constDefaultViewNavigationFGColor = "000000"
  const constDefaultGridHorizontalStripeBGColor = "#f7f3f7"
  const constDefaultGridUnselectedIndexFGColor = "#000000"
  const constDefaultGridSelectedIndexBGColor = "#000000"
  const constDefaultGridSelectedIndexFGColor = "#ffffff"
  const constDefaultGlobalTableBGColor = "#ffffff"
  const constDefaultGlobalTableFGColor = "#000000"
  const constDefaultGlobalHeaderBGColor = "#e9e8e4"
  const constDefaultGlobalHeaderFGColor = "blue"

  const constVarTypeBoolean = 11
%>