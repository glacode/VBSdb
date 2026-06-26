<%
'***********************************************
function funIntVbsDbFilterFieldType( strFilterFieldName , objVbsDb )
'***********************************************
  funIntVbsDbFilterFieldType = _
    cInt( request.form( funStrVbsDbHiddenFieldTypeName( strFilterFieldName , objVbsDb ) ) )
end function

'***********************************************
function funStrFieldWhereClauseWithFieldValue( intFieldType , strFieldAlias , strFieldName , strFieldValue , objVbsDb )
'***********************************************
  select case intFieldType
  	case vbsDbAdGUID
  	  ' numeric
      funStrFieldWhereClauseWithFieldValue = funStrFieldWhereClauseForFieldOfTypeAdGUID( strFieldName , strFieldValue , objVbsDb )
  	case vbsDbAdSmallInt, vbsDbAdInteger, vbsDbAdDecimal, vbsDbAdSingle, vbsDbAdDouble, vbsDbAdCurrency, _
  	  vbsDbAdTinyInt, vbsDbAdUnsignedTinyInt, vbsDbAdUnsignedSmallInt, vbsDbAdUnsignedInt , vbsDbAdBigInt, vbsDbAdUnsignedBigInt, _
  	  vbsDbAdNumeric, vbsDbAdVarNumeric
  	  ' numeric
  			if funBolGlobalBooleanField( strFieldName , objVbsDb ) then
  				' the field is defined boolean using the GlobalBooleanFields property
					funStrFieldWhereClauseWithFieldValue = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & "=" & funStrGlobalDbTypeBooleanFieldFormattedValue( strFieldValue , objVbsDb )
  			else
  				' the field is not defined boolean using the GlobalBooleanFields property
					funStrFieldWhereClauseWithFieldValue = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & _
						funStrVbsDbSearchOperator( strFieldAlias , objVbsDb ) & strFieldValue
				end if
  	case vbsDbAdBoolean	
  	  ' boolean
       funStrFieldWhereClauseWithFieldValue = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & "=" & funStrGlobalDbTypeBooleanFieldFormattedValue( strFieldValue , objVbsDb )
   	case vbsDbAdChar, vbsDbAdWChar, vbsDbAdVarChar, vbsDbAdLongVarChar, vbsDbAdVarWChar, vbsDbAdLongVarWChar, vbsDbAdVarBinary
  	  ' strings
  	  funStrFieldWhereClauseWithFieldValue = _
  	    funStrGlobalDbTypeStringFieldWhereClauseWithFieldValue( strFieldName , strFieldValue , objVbsDb )
  	case vbsDbAdDate, vbsDbAdDBDate, vbsDbAdDBTime, vbsDbAdDBTimeStamp
  	  ' dates/time
  	  'funStrFieldWhereClause = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & "=#" & month( strFieldValue ) & "/" & day( strFieldValue ) & "/" & year( strFieldValue ) & " " & hour( strFieldValue ) & "." & minute( strFieldValue ) & "." & second( strFieldValue ) & "#"
  	  funStrFieldWhereClauseWithFieldValue = funStrGlobalDbTypeFieldName( strFieldName , objVbsDb ) & _
  	    funStrVbsDbSearchOperator( strFieldAlias , objVbsDb ) & funStrGlobalDbTypeDateFieldFormattedValue( strFieldValue , objVbsDb )
  	case else
  	  drawError "Function funStrFieldWhereClause: type " &_
  	    intFieldType & " is not covered. The field name is " & strFieldName & _
  	    ". Please, signal this bug message to <a href='mailto:support@vbsdb.com'>support@vbsdb.com</a>." , objVbsDb
  end select
end function

'***********************************************
function funStrFilterFieldWhereClauseActually( strFormControlName , strFilterFieldName , objVbsDb )
'***********************************************
  funStrFilterFieldWhereClauseActually = funStrFieldWhereClauseWithFieldValue( _
    funIntVbsDbFilterFieldType( strFilterFieldName , objVbsDb ) , _
    strFilterFieldName , _
    funStrSearchAliasField( strFilterFieldName , objVbsDb ) , _
    request.form( strFormControlName ) , objVbsDb )
end function

'***********************************************
function funStrFilterFieldWhereClauseNotEmpty( strFormControlName , strFilterFieldName , objVbsDb )
'***********************************************
  funStrFilterFieldWhereClauseNotEmpty = "(" & funStrFilterFieldWhereClauseActually( strFormControlName , strFilterFieldName , objVbsDb ) & ")" & " and "
end function

'***********************************************
function funStrFilterFieldWhereClause( strFormControlName , strFilterFieldName , objVbsDb )
'***********************************************
  if trim( request.form( strFormControlName ) ) <> "" then
    ' the user inserted some value for this filter field
    funStrFilterFieldWhereClause = funStrFilterFieldWhereClauseNotEmpty( strFormControlName , strFilterFieldName , objVbsDb )
  else
    ' the user inserted no value for this filter field
    funStrFilterFieldWhereClause = ""
  end if
end function

'***********************************************
function funStrFilterWhereClauseListItemActually( strFormControlName , objVbsDb )
'***********************************************
  dim strFilterFieldName
  strFilterFieldName = funStrFilterFieldName( strFormControlName )
  funStrFilterWhereClauseListItemActually = funStrFilterFieldWhereClause( strFormControlName , strFilterFieldName , objVbsDb )
end function

'***********************************************
function funStrFilterWhereClauseListItem( strFormControlName , objVbsDb )
'***********************************************
  if funBolVbsDbIsFilterField( strFormControlName , objVbsDb ) then
    ' the current form control is a filter field control (it is not an hidden form control)
    funStrFilterWhereClauseListItem = funStrFilterWhereClauseListItemActually( strFormControlName , objVbsDb )
  else
    ' the current form control is not a filter field control (probably it is an hidden form controll, added by vspDb procedures)
    funStrFilterWhereClauseListItem = ""
  end if
end function

'***********************************************
function funStrFilterWhereClauseList( objVbsDb )
'***********************************************
  dim strFormControlName
  funStrFilterWhereClauseList = ""
  for each strFormControlName in Request.Form
    funStrFilterWhereClauseList = funStrFilterWhereClauseList & funStrFilterWhereClauseListItem( strFormControlName , objVbsDb )
  next
  funStrFilterWhereClauseList = cutFinal( funStrFilterWhereClauseList , " and " )
end function

'***********************************************
sub vbsDbSetFilterWhereClauseActually( objVbsDb )
'***********************************************
  dim strFilterWhereClauseList
  strFilterWhereClauseList = funStrFilterWhereClauseList( objVbsDb )
  if strFilterWhereClauseList <> "" then
    ' the user defined some filter clause
    objVbsDb( "FilterWhereClause" ) = "(" & strFilterWhereClauseList & ")"
  else
    ' the user didn't define some filter clause (he/she wants to reset the filter where clause)
    objVbsDb( "FilterWhereClause" ) = ""
  end if    
end sub

'***********************************************
sub vbsDbRemoveFilterWhereClause( objVbsDb )
'***********************************************
  objVbsDb( "FilterWhereClause" ) = ""
end sub

'***********************************************
sub vbsDbSetFilterWhereClause( byRef objVbsDb )
'***********************************************
  if objVbsDb( "RequestVBSdbClickClass" ) = "VBSdbApplyFilter" then
    ' the screen whe are coming from was a filter screen
    vbsDbCheckSearchAliasFieldsErrors objVbsDb
    vbsDbSetFilterWhereClauseActually objVbsDb
  elseif VbsDbGetLastAction( objVbsDb ) = "RemoveFilter" then
    vbsDbRemoveFilterWhereClause objVbsDb
  end if
end sub

'***********************************************
'function funIntFromBegin( objVbsDb )
'***********************************************
'  funIntFromBegin = inStr( 1 , objVbsDb( "Sql" ) , " FROM " , 1 )
'end function

'***********************************************
'function funIntSqlClauseBegin( strClauseReservedName , objVbsDb )
'***********************************************
'  dim intFromBegin , intClauseBegin
'  intFromBegin = funIntFromBegin( objVbsDb )
'  intClauseBegin = inStr( intFromBegin , objVbsDb( "Sql" ) , strClauseReservedName , 1 )
'  if intClauseBegin < intFromBegin then
    ' the sql clause is not present
'    funIntSqlClauseBegin = len( objVbsDb( "Sql" ) ) + 1
'  else
    ' the sql clause is present
'    funIntSqlClauseBegin = intClauseBegin
'  end if
'end function

'***********************************************
'function funIntForBrowseBegin( objVbsDb )
'***********************************************
'  funIntForBrowseBegin = funIntSqlClauseBegin( " FOR " , objVbsDb )
'end function

'***********************************************
'function funIntComputeBegin( objVbsDb )
'***********************************************
'  funIntComputeBegin = funIntSqlClauseBegin( " COMPUTE " , objVbsDb )
'end function

'***********************************************
'function funIntOrderByBegin( objVbsDb )
'***********************************************
'  funIntOrderByBegin = funIntSqlClauseBegin( " ORDER " , objVbsDb )
'end function

'***********************************************
'function funIntHavingBegin( objVbsDb )
'***********************************************
'  funIntHavingBegin = funIntSqlClauseBegin( " HAVING " , objVbsDb )
'end function

'***********************************************
'function funIntGroupByBegin( objVbsDb )
'***********************************************
'  funIntGroupByBegin = funIntSqlClauseBegin( " GROUP " , objVbsDb )
'end function

'***********************************************
'function funIntFinalSemicolon( objVbsDb )
'***********************************************
'  if right( objVbsDb( "Sql" ) , 1 ) = ";" then
    ' the last Sql character is a semicolon
'    funIntFinalSemicolon = len( objVbsDb( "Sql" ) )
'  else
    ' the last Sql character is not a semicolon
'    funIntFinalSemicolon = len( objVbsDb( "Sql" ) ) + 1
'  end if
'end function

'***********************************************
'function funIntWhereClauseEnd( objVbsDb )
'***********************************************
'  dim intFromBegin
'  funIntWhereClauseEnd = funIntForBrowseBegin( objVbsDb ) - 1
'  funIntWhereClauseEnd = funNumMin( funIntWhereClauseEnd , funIntComputeBegin( objVbsDb ) - 1 )
'  funIntWhereClauseEnd = funNumMin( funIntWhereClauseEnd , funIntOrderByBegin( objVbsDb ) - 1 )
'  funIntWhereClauseEnd = funNumMin( funIntWhereClauseEnd , funIntHavingBegin( objVbsDb ) - 1 )
'  funIntWhereClauseEnd = funNumMin( funIntWhereClauseEnd , funIntGroupByBegin( objVbsDb ) - 1 )
'  funIntWhereClauseEnd = funNumMin( funIntWhereClauseEnd , funIntFinalSemicolon( objVbsDb ) - 1 )
'end function

'***********************************************
'function funStrSqlWithFilterWhereClauseBegin( objVbsDb )
'***********************************************
'  funStrSqlWithFilterWhereClauseBegin = _
'    left( objVbsDb( "Sql" ) , funIntWhereClauseEnd( objVbsDb ) )
'end function

'***********************************************
'function funBolWhereClauseIsAlreadyInSql( objVbsDb )
'***********************************************
'  dim intFromBegin , intClauseBegin
'  intFromBegin = funIntFromBegin( objVbsDb )
'  intClauseBegin = inStr( intFromBegin , objVbsDb( "Sql" ) , " WHERE " , 1 )
'  funBolWhereClauseIsAlreadyInSql = ( intFromBegin < intClauseBegin )
'end function

'***********************************************
'function funStrSqlWithFilterWhereClauseBody( objVbsDb )
'***********************************************
'  if funBolWhereClauseIsAlreadyInSql( objVbsDb ) then
    ' the Sql property already contains a where clause
'    funStrSqlWithFilterWhereClauseBody = " and " & objVbsDb( "FilterWhereClause" )
'  else
    ' the Sql property doesn't contain any where clause
'    funStrSqlWithFilterWhereClauseBody = " where " & objVbsDb( "FilterWhereClause" )
'  end if
'end function

'***********************************************
'function funStrSqlWithFilterWhereClauseEnd( objVbsDb )
'***********************************************
'  funStrSqlWithFilterWhereClauseEnd = _
'    mid( objVbsDb( "Sql" ) , funIntWhereClauseEnd( objVbsDb ) + 1 )
'end function

'***********************************************
'function funStrSqlWithFilterWhereClause( objVbsDb )
'***********************************************
'  funStrSqlWithFilterWhereClause = funStrSqlWithFilterWhereClauseBegin( objVbsDb ) & _
'    funStrSqlWithFilterWhereClauseBody( objVbsDb ) & funStrSqlWithFilterWhereClauseEnd( objVbsDb )
'end function

'***********************************************
'***********************************************
sub vbsDbSetSqlWithFilterWhereClause( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetFilterWhereClause objVbsDb
  if objVbsDb( "FilterWhereClause" ) = "" then
    ' the user didn't define a filter
    objVbsDb( "SqlWithFilterWhereClause" ) = objVbsDb( "Sql" )
  else
    ' the user defined a nonempty filter
    'objVbsDb( "SqlWithFilterWhereClause" ) = funStrSqlWithFilterWhereClause( objVbsDb )
    objVbsDb( "SqlWithFilterWhereClause" ) = _
      funStrSqlInsertWhereClause( objVbsDb( "Sql" ) , objVbsDb( "FilterWhereClause" ) , objVbsDb )
  end if
end sub
%>