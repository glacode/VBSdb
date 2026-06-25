<%
'***********************************************
function funStrLeftPadZero( intToLeftPad , intTotalChars )
'***********************************************
  dim intCurrChar
  funStrLeftPadZero = ""
  for intCurrChar = 1 to intTotalChars - len( cStr( intToLeftPad ) )
  	funStrLeftPadZero = funStrLeftPadZero & "0"
  next
  funStrLeftPadZero = funStrLeftPadZero & cStr( intToLeftPad )
end function

'***********************************************
function funStrYear( strControlValue , objVbsDb )
'***********************************************
  funStrYear = funStrLeftPadZero( year( strControlValue ) , 4 )
end function

'***********************************************
function funStrMonth( strControlValue , objVbsDb )
'***********************************************
  funStrMonth = funStrLeftPadZero( month( strControlValue ) , 2 )
end function

'***********************************************
function funStrDay( strControlValue , objVbsDb )
'***********************************************
  funStrDay = funStrLeftPadZero( day( strControlValue ) , 2 )
end function

'***********************************************
function funStrHour( strControlValue , objVbsDb )
'***********************************************
  funStrHour = funStrLeftPadZero( hour( strControlValue ) , 2 )
end function

'***********************************************
function funStrMinute( strControlValue , objVbsDb )
'***********************************************
  funStrMinute = funStrLeftPadZero( minute( strControlValue ) , 2 )
end function

'***********************************************
function funStrSecond( strControlValue , objVbsDb )
'***********************************************
  funStrSecond = funStrLeftPadZero( second( strControlValue ) , 2 )
end function

'***********************************************
'***********************************************
function funStrGlobalDbTypeStringFieldWhereClauseWithFieldValue( strFieldName , strFieldValue , objVbsDb )
'***********************************************
'***********************************************
  dim strLikePart
  strLikePart = " like '%" & replace( uCase( strFieldValue ) , "'" , "''" ) & "%'"
  if objVbsDb( "GlobalDbType" ) = "FOXPRO" then
    funStrGlobalDbTypeStringFieldWhereClauseWithFieldValue = _
      "UPPER(" & funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & ") " & _
      strLikePart 
  else
    funStrGlobalDbTypeStringFieldWhereClauseWithFieldValue = _
      funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & _
      strLikePart
  end if
end function

'***********************************************
'***********************************************
function funStrGlobalDbTypeSelectTop1( objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "GlobalDbType" ) = "ACCESS" then
    funStrGlobalDbTypeSelectTop1 = "top 1 "
  else
    funStrGlobalDbTypeSelectTop1 = ""
  end if
end function

'***********************************************
function funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForAccess( strControlValue , objVbsDb )
'***********************************************
  funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForAccess = _
    "#" & funStrMonth( strControlValue , objVbsDb ) & "/" & funStrDay( strControlValue , objVbsDb ) & "/" & funStrYear( strControlValue , objVbsDb ) & " " & _
    funStrHour( strControlValue , objVbsDb ) & "." & funStrMinute( strControlValue , objVbsDb ) & "." & funStrSecond( strControlValue , objVbsDb ) & "#"
end function

'***********************************************
function funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForFoxpro( strControlValue , objVbsDb )
'***********************************************
  funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForFoxpro = _
    "{" & funStrMonth( strControlValue , objVbsDb ) & "/" & funStrDay( strControlValue , objVbsDb ) & "/" & funStrYear( strControlValue , objVbsDb ) & " " & _
    funStrHour( strControlValue , objVbsDb ) & "." & funStrMinute( strControlValue , objVbsDb ) & "." & funStrSecond( strControlValue , objVbsDb ) & "}"
end function

'***********************************************
function funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForOracle( strControlValue , objVbsDb )
'***********************************************
  funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForOracle = _
    "to_date('" & funStrYear( strControlValue , objVbsDb ) & "-" & funStrMonth( strControlValue , objVbsDb ) & "-" & funStrDay( strControlValue , objVbsDb ) & _
    " " & funStrHour( strControlValue , objVbsDb ) & ":" & funStrMinute( strControlValue , objVbsDb ) & ":" & funStrSecond( strControlValue , objVbsDb ) & _
    "','YYYY-MM-DD HH24:MI:SS')"
end function

'***********************************************
function funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForSqlServer( strControlValue , objVbsDb )
'***********************************************
  funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForSqlServer = _
    "'" & funStrMonth( strControlValue , objVbsDb ) & "/" & funStrDay( strControlValue , objVbsDb ) & "/" & funStrYear( strControlValue , objVbsDb ) & " " & _
    funStrHour( strControlValue , objVbsDb ) & ":" & funStrMinute( strControlValue , objVbsDb ) & ":" & funStrSecond( strControlValue , objVbsDb ) & "'"
end function

'***********************************************
function funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForMySql( strControlValue , objVbsDb )
'***********************************************
  funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForMySql = _
    "'" & funStrYear( strControlValue , objVbsDb ) & "-" & funStrMonth( strControlValue , objVbsDb ) & "-" & funStrDay( strControlValue , objVbsDb ) & "'"
end function

'***********************************************
function funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForOtherDbTypes( strControlValue , objVbsDb )
'***********************************************
  funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForOtherDbTypes = _
    "'" & funStrMonth( strControlValue , objVbsDb ) & "/" & funStrDay( strControlValue , objVbsDb ) & "/" & funStrYear( strControlValue , objVbsDb ) & " " & _
    funStrHour( strControlValue , objVbsDb ) & "." & funStrMinute( strControlValue , objVbsDb ) & "." & funStrSecond( strControlValue , objVbsDb ) & "'"
end function

'***********************************************
function funStrGlobalDbTypeDateFieldFormattedValueNonEmptyValidDate( strControlValue , objVbsDb )
'***********************************************
  select case objVbsDb( "GlobalDbType" )
    case "ACCESS"
      funStrGlobalDbTypeDateFieldFormattedValueNonEmptyValidDate = _
        funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForAccess( strControlValue , objVbsDb )
    case "FOXPRO"
      funStrGlobalDbTypeDateFieldFormattedValueNonEmptyValidDate = _
        funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForFoxpro( strControlValue , objVbsDb )
    case "ORACLE"
      funStrGlobalDbTypeDateFieldFormattedValueNonEmptyValidDate = _
        funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForOracle( strControlValue , objVbsDb )
    case "SQLSERVER"
      funStrGlobalDbTypeDateFieldFormattedValueNonEmptyValidDate = _
        funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForSqlServer( strControlValue , objVbsDb )
    case "MYSQL"
      funStrGlobalDbTypeDateFieldFormattedValueNonEmptyValidDate = _
        funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForMySql( strControlValue , objVbsDb )
    case else
      funStrGlobalDbTypeDateFieldFormattedValueNonEmptyValidDate = _
        funStrGlobalDbTypeDateFieldFormattedValueNonEmptyForOtherDbTypes( strControlValue , objVbsDb )
  end select
end function

'***********************************************
function funStrGlobalDbTypeDateFieldFormattedValueNonEmpty( strControlValue , objVbsDb )
'***********************************************
  dim strErrorMessage
  if isDate( strControlValue ) then
    ' the user inserted a valid date value in the input control
    funStrGlobalDbTypeDateFieldFormattedValueNonEmpty = _
      funStrGlobalDbTypeDateFieldFormattedValueNonEmptyValidDate( strControlValue , objVbsDb )
  else
    ' the user did not inserted a valid date value in the input control
    strErrorMessage = "The user inserted the following invalid date value:<br><br>" & _
      strControlValue & "<br><br>" & _
      "Use the InputValidateDateFields property to intercept such invalid input errors on the " & _
      "client side, by means of a Javascript validation function."
    'drawErrorWithDebugging "InputSelectFields is assigned a wrong select statement: " & strSql  , objVbsDb
    drawError strErrorMessage , objVbsDb
  end if
end function

'***********************************************
'***********************************************
function funStrGlobalDbTypeDateFieldFormattedValue( strControlValue , objVbsDb )
'***********************************************
'***********************************************
  if isNull( strControlValue ) or trim( strControlValue ) = "" then
    ' date field is null or empty
    funStrGlobalDbTypeDateFieldFormattedValue = "null"
  else
    funStrGlobalDbTypeDateFieldFormattedValue = _
      funStrGlobalDbTypeDateFieldFormattedValueNonEmpty( strControlValue , objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funStrGlobalDbTypeTrueBooleanFieldFormattedValue( objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "GlobalDbType" ) = "SQLSERVER" or objVbsDb( "GlobalDbType" ) = "MYSQL" then
    funStrGlobalDbTypeTrueBooleanFieldFormattedValue = "1"
  else
    funStrGlobalDbTypeTrueBooleanFieldFormattedValue = "true"
  end if
end function

'***********************************************
'***********************************************
function funStrGlobalDbTypeFalseBooleanFieldFormattedValue( objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "GlobalDbType" ) = "SQLSERVER" or objVbsDb( "GlobalDbType" ) = "MYSQL" then
    funStrGlobalDbTypeFalseBooleanFieldFormattedValue = "0"
  else
    funStrGlobalDbTypeFalseBooleanFieldFormattedValue = "false"
  end if
end function

'***********************************************
'***********************************************
function funStrGlobalDbTypeBooleanFieldFormattedValue( strControlValue , objVbsDb )
'***********************************************
'***********************************************
  if ( ( uCase( strControlValue ) ="TRUE" ) or ( uCase( strControlValue ) ="ON" ) or ( uCase( strControlValue ) ="YES" ) or ( uCase( strControlValue ) ="SI" ) or ( strControlValue = "1" ) ) then
  	'funStrDbEditBooleanFieldFormattedValue = "true"
  	funStrGlobalDbTypeBooleanFieldFormattedValue = funStrGlobalDbTypeTrueBooleanFieldFormattedValue( objVbsDb )
  else
  	'funStrDbEditBooleanFieldFormattedValue = "false"
  	funStrGlobalDbTypeBooleanFieldFormattedValue = funStrGlobalDbTypeFalseBooleanFieldFormattedValue( objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funStrGlobalDbTypeEditTableName( objVbsDb )
'***********************************************
'***********************************************
  if inStr( objVbsDb( "EditTableName" ) , " " ) > 0 then
    ' EditTableName contains spaces
    funStrGlobalDbTypeEditTableName = "[" & objVbsDb( "EditTableName" ) & "]"
  else
    ' EditTableName does not contain spaces
    funStrGlobalDbTypeEditTableName = objVbsDb( "EditTableName" )
  end if
end function

'***********************************************
function funStrGlobalDbTypeFieldNameAtomic( strFieldName , objVbsDb )
'***********************************************
  if ( inStr( strFieldName , " " ) > 0 ) and ( inStr( strFieldName , "[" ) = 0 ) then
    ' strFieldName contains spaces
    funStrGlobalDbTypeFieldNameAtomic = "[" & strFieldName & "]"
  else
    ' strFieldName does not contain spaces
    funStrGlobalDbTypeFieldNameAtomic = strFieldName
  end if
end function

'***********************************************
'***********************************************
function funStrGlobalDbTypeFieldName( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  dim intDotPosition
  intDotPosition = inStr( strFieldName , "." )
  if intDotPosition > 0 then
    ' strFieldName contains a dot: it is in the form tableName.fieldName
    funStrGlobalDbTypeFieldName = _
      funStrGlobalDbTypeFieldNameAtomic( left( strFieldName , intDotPosition - 1 ) , objVbsDb ) & _
      "." & funStrGlobalDbTypeFieldNameAtomic( mid( strFieldName , intDotPosition + 1 ) , objVbsDb )
  else
    ' strFieldName doesn't contains a dot: the field table name is not specified
    funStrGlobalDbTypeFieldName = _
      funStrGlobalDbTypeFieldNameAtomic( strFieldName , objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funIntCheckAdoxFieldTypeForOracleDatabase( funIntAdoxFieldType , objVbsDb )
'***********************************************
'***********************************************
  if ( ( objVbsDb( "GlobalDbType" ) = "ORACLE" ) and ( funIntAdoxFieldType = -1 ) ) then
    ' current database if of Oracle type and field type is numeric (adox mistakely treates Oracle numeric fields; I believe it to be an Adox bug)
    funIntCheckAdoxFieldTypeForOracleDatabase = 139
  else
    ' current database is not of Oracle type or field type is not numeric
    funIntCheckAdoxFieldTypeForOracleDatabase = funIntAdoxFieldType
  end if
end function
%>