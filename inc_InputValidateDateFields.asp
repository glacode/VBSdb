<%
'***********************************************
'function funStrInputValidateDateFormat()
'***********************************************
'  select case session.lcid
'    case 1027 , 1034 , 1036 , 1040 , 1046 , 1057 , 1086 , 2057 , 2058 , 2060 , 2067 , 2110 , 3081 , 3082 , 4105 , 4106 , 5129 , 5130 , 5132 , 6153 , 6156 , 7178 , 8201 , 8202 , 9226 , 10249 , 10250 , 11273 , 11274 , 12298 , 14346 , 15370 , 16394 , 17418 , 18442 , 19466 , 20490
'      funStrInputValidateDateFormat = "DD/MM/YYYY hh:mm:ss"
'    case 2048
'      funStrInputValidateDateFormat = "DD/MM/YYYY hh.mm.ss"
'    case 1030 , 1043 , 2070 , 13322
'      funStrInputValidateDateFormat = "DD-MM-YYYY hh:mm:ss"
'    case 1080
'      funStrInputValidateDateFormat = "DD-MM-YYYY hh.mm.ss"
'    case 1031 , 1035 , 1039 , 1044 , 2055 , 2064 , 2068 , 2077 , 3079 , 4103 , 4108 , 5127 , 66567
'      funStrInputValidateDateFormat = "DD.MM.YYYY hh:mm:ss"
'    case 1033 , 1089 , 6154 , 9225 , 12297 , 13321
'      funStrInputValidateDateFormat = "MM/DD/YYYY hh:mm:ss"
'    case 1053 , 3084
'      funStrInputValidateDateFormat = "YYYY-MM-DD hh:mm:ss"
'    case 1069 , 1078 , 7177
'      funStrInputValidateDateFormat = "YYYY/MM/DD hh:mm:ss"
'  end select
'end function

'***********************************************
function funIntDateFormatPositionForDatePart( strDateItemIdentifier , objVbsDb )
'***********************************************
  select case inStr( VbsDbGetDateFormat( objVbsDb ) , strDateItemIdentifier )
    case 1
      funIntDateFormatPositionForDatePart = 1
    case 4 , 6
      funIntDateFormatPositionForDatePart = 2
    case else
      funIntDateFormatPositionForDatePart = 3
  end select
end function

'***********************************************
function funIntDateFormatPosition( strDateItemIdentifier , objVbsDb )
'***********************************************
  select case strDateItemIdentifier
    case "hh"
      funIntDateFormatPosition = 5
    case "mm"
      funIntDateFormatPosition = 6
    case "ss"
      funIntDateFormatPosition = 7
    case else
      ' strDateItemIdentifier is either "YYYY" or "MM" or "DD"
      funIntDateFormatPosition = funIntDateFormatPositionForDatePart( strDateItemIdentifier , objVbsDb )
  end select
end function

'***********************************************
sub vbsDbSetDictionaryInputValidateDateFieldsHandleSingleQuote( byRef objVbsDb )
'***********************************************
  dim strFieldName
  for each strFieldName in objVbsDb( "DictionaryInputValidateDateFields" )
    objVbsDb( "DictionaryInputValidateDateFields" )( strFieldName ) = _
      replace( objVbsDb( "DictionaryInputValidateDateFields" )( strFieldName ) , "'" , "\'" )
  next
end sub

'***********************************************
sub vbsDbSetDictionaryInputValidateDateFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "InputValidateDateFields" , objVbsDb
  vbsDbSetDictionaryInputValidateDateFieldsHandleSingleQuote( objVbsDb )
end sub

'***********************************************
'***********************************************
sub vbsDbSetInputValidateDateRelatedProperties( objVbsDb )
'***********************************************
'***********************************************
  'vbsDbSetDictionaryInputValidateDateFormatRelatedProperties objVbsDb
  vbsDbSetDictionaryInputValidateDateFields objVbsDb
end sub

'***********************************************
sub drawValidateDateFunctionCallLineActually( strFieldName , objVbsDb )
'***********************************************
%>
    message = message + vbsDbValidateDate(form.<%=funStrFormFieldName( strFieldName , objVbsDb )%>,"<%=objVbsDb( "DictionaryInputValidateDateFields" )( strFieldName )%>")
<%
end sub

'***********************************************
sub drawValidateDateFunctionCallLine( strFieldName , objVbsDb )
'***********************************************
  if objVbsDb( "DictionaryInputFields" ).exists( strFieldName ) and not funBolEditFieldIsReadOnly( strFieldName , objVbsDb ) then
    ' strFieldName is an input field in the current screen and it is to be validated as a date field
    drawValidateDateFunctionCallLineActually strFieldName , objVbsDb
  end if
end sub

'***********************************************
'***********************************************
sub drawValidateFunctionCallLinesForInputValidateDateFields( objVbsDb )
'***********************************************
'***********************************************
  dim strFieldName
  for each strFieldName in objVbsDb( "DictionaryInputValidateDateFields" )
    drawValidateDateFunctionCallLine strFieldName , objVbsDb
  next
end sub

'***********************************************
function funStrRegularExpressionForDateSeparators( objVbsDb )
'***********************************************
  funStrRegularExpressionForDateSeparators = "[" & funStrValidDateSeparators( objVbsDb ) & "]"
  funStrRegularExpressionForDateSeparators = _
    replace( funStrRegularExpressionForDateSeparators , "/" , "\/" )
  funStrRegularExpressionForDateSeparators = _
    replace( funStrRegularExpressionForDateSeparators , "." , "\." )
end function

'***********************************************
function funStrDateDateFormatWithReplacedDateSeparator( objVbsDb )
'***********************************************
  funStrDateDateFormatWithReplacedDateSeparator = left( objVbsDb( "GlobalDateFormat" ) , 10 )
  funStrDateDateFormatWithReplacedDateSeparator = _
    replace( funStrDateDateFormatWithReplacedDateSeparator , _
    funStrStandardDateSeparator( objVbsDb ) , funStrRegularExpressionForDateSeparators( objVbsDb ) )
end function

'***********************************************
function funStrRegularExpressionForTimeSeparators( objVbsDb )
'***********************************************
  funStrRegularExpressionForTimeSeparators = "[" & funStrValidTimeSeparators( objVbsDb ) & "]"
  funStrRegularExpressionForTimeSeparators = _
    replace( funStrRegularExpressionForTimeSeparators , "/" , "\/" )
  funStrRegularExpressionForTimeSeparators = _
    replace( funStrRegularExpressionForTimeSeparators , "." , "\." )
end function

'***********************************************
function funStrTimeDateFormatWithReplacedTimeSeparator( objVbsDb )
'***********************************************
  funStrTimeDateFormatWithReplacedTimeSeparator = mid( objVbsDb( "GlobalDateFormat" ) , 12 )
  funStrTimeDateFormatWithReplacedTimeSeparator = _
    replace( funStrTimeDateFormatWithReplacedTimeSeparator , _
    funStrStandardTimeSeparator( objVbsDb ) , funStrRegularExpressionForTimeSeparators( objVbsDb ) )
end function

'***********************************************
function funStrDateFormatRegularExpression( objVbsDb )  ' returns the regular expression corresponding to the chosen date format
'***********************************************
  funStrDateFormatRegularExpression = VbsDbGetDateFormat( objVbsDb )
  funStrDateFormatRegularExpression = _
    funStrDateDateFormatWithReplacedDateSeparator( objVbsDb ) & " " & _
    funStrTimeDateFormatWithReplacedTimeSeparator( objVbsDb )
  funStrDateFormatRegularExpression = replace( funStrDateFormatRegularExpression , " " , "[ ]*" )
  funStrDateFormatRegularExpression = "[ ]*\b" & funStrDateFormatRegularExpression & ""
  funStrDateFormatRegularExpression = replace( funStrDateFormatRegularExpression , "YYYY" , "(\d{4})" )
  funStrDateFormatRegularExpression = replace( funStrDateFormatRegularExpression , "MM" , "(\d{1,2})" )
  funStrDateFormatRegularExpression = replace( funStrDateFormatRegularExpression , "DD" , "(\d{1,2})" )
  funStrDateFormatRegularExpression = replace( funStrDateFormatRegularExpression , "hh" , "\b((\d{1,2})" )
  funStrDateFormatRegularExpression = replace( funStrDateFormatRegularExpression , "mm" , "(\d{1,2})" )
  funStrDateFormatRegularExpression = replace( funStrDateFormatRegularExpression , "ss" , "(\d{1,2}))*" )
  funStrDateFormatRegularExpression = funStrDateFormatRegularExpression & "([ ]{1,}(PM|p.m.|AM|a.m.)){0,1}"
end function

'***********************************************
'***********************************************
sub drawJavascriptFunValidateDate( objVbsDb )
'***********************************************
'***********************************************
%>
  ////////////////////////
  function funIntRegExpYear( arrRegExpItems )
  ////////////////////////
  {
    return arrRegExpItems[ <%=funIntDateFormatPosition( "YYYY" , objVbsDb )%> ];
  }
  
  ////////////////////////
  function funBolIsGoodDateWithRegExpYearOk( arrRegExpItems )
  ////////////////////////
  {
    var intYear = funIntRegExpYear( arrRegExpItems );
    return ( ( intYear >= 0 ) && ( intYear <= 9999 ) );
  }

  ////////////////////////
  function funIntRegExpMonth( arrRegExpItems )
  ////////////////////////
  {
    return Number( arrRegExpItems[ <%=funIntDateFormatPosition( "MM" , objVbsDb )%> ] );
  }
  
  ////////////////////////
  function funBolIsGoodDateWithRegExpMonthOk( arrRegExpItems )
  ////////////////////////
  {
    var intMonth = funIntRegExpMonth( arrRegExpItems );
    return ( ( intMonth >= 1 ) && ( intMonth <= 12 ) );
  }

  ////////////////////////
  function funIntRegExpDay( arrRegExpItems )
  ////////////////////////
  {
    return arrRegExpItems[ <%=funIntDateFormatPosition( "DD" , objVbsDb )%> ];
  }
  
  ////////////////////////
  function funBolIsGood29FebruaryDayActually( intYear )
  ////////////////////////
  {
    return ( ( ( intYear % 1000 ) == 0 ) || ( ( ( intYear % 4 ) == 0 ) && !( ( intYear % 100 ) == 0 ) ) );
  }
  
  ////////////////////////
  function funBolIsGood29FebruaryDay( intDay , intYear )
  ////////////////////////
  {
    return ( ( intDay == 29 ) && funBolIsGood29FebruaryDayActually( intYear ) );
  }
  
  ////////////////////////
  function funBolIsGoodFebruaryDay( intDay , intYear )
  ////////////////////////
  {
    return ( ( intDay <= 28 ) || funBolIsGood29FebruaryDay( intDay , intYear ) );
  }
  
  ////////////////////////
  function funBolIsGoodDateWithRegExpDayOkWithDayMonthAndYear( intDay , intMonth , intYear )
  ////////////////////////
  {
    var bolIsGoodDay
    bolIsGoodDay = ( intDay >= 1 )
    if ( bolIsGoodDay )
      switch ( intMonth ) {
        case 1:
          bolIsGoodDay = ( intDay <= 31 );
          break;
        case 2:
          bolIsGoodDay = funBolIsGoodFebruaryDay( intDay , intYear );
          break;
        case 3:
          bolIsGoodDay = ( intDay <= 31 );
          break;
        case 4:
          bolIsGoodDay = ( intDay <= 30 );
          break;
        case 5:
          bolIsGoodDay = ( intDay <= 31 );
          break;
        case 6:
          bolIsGoodDay = ( intDay <= 30 );
          break;
        case 7:
          bolIsGoodDay = ( intDay <= 31 );
          break;
        case 8:
          bolIsGoodDay = ( intDay <= 31 );
          break;
        case 9:
          bolIsGoodDay = ( intDay <= 30 );
          break;
        case 10:
          bolIsGoodDay = ( intDay <= 31 );
          break;
        case 11:
          bolIsGoodDay = ( intDay <= 30 );
          break;
        case 12:
          bolIsGoodDay = ( intDay <= 31 );
          break;
    }
    return bolIsGoodDay;
  }
  
  ////////////////////////
  function funBolIsGoodDateWithRegExpDayOk( arrRegExpItems )
  ////////////////////////
  {
    var intDay = funIntRegExpDay( arrRegExpItems );
    var intMonth = funIntRegExpMonth( arrRegExpItems );
    var intYear = funIntRegExpYear( arrRegExpItems );
    return funBolIsGoodDateWithRegExpDayOkWithDayMonthAndYear( intDay , intMonth , intYear );
  }

  ////////////////////////
  function funBolHourIsInserted( arrRegExpItems )
  ////////////////////////
  {
    return ( arrRegExpItems.length >= ( <%=funIntDateFormatPosition( "hh" , objVbsDb )%> + 1 ) );
  }
  
  ////////////////////////
  function funIntRegExpHour( arrRegExpItems )
  ////////////////////////
  {
    return arrRegExpItems[ <%=funIntDateFormatPosition( "hh" , objVbsDb )%> ];
  }
  
  ////////////////////////
  function funBolIsGoodDateWithRegExpHourOk( arrRegExpItems )
  ////////////////////////
  {
    if ( funBolHourIsInserted( arrRegExpItems ) )
    {
      // the user inserted hours
      var intHour = funIntRegExpHour( arrRegExpItems );
      return ( ( intHour >= 0 ) && ( intHour <= 23 ) );
    }
    else
      // the user didn't insert hours
      return true;
  }

  ////////////////////////
  function funBolMinuteIsInserted( arrRegExpItems )
  ////////////////////////
  {
    return ( arrRegExpItems.length >= ( <%=funIntDateFormatPosition( "mm" , objVbsDb )%> + 1 ) );
  }

  ////////////////////////
  function funIntRegExpMinute( arrRegExpItems )
  ////////////////////////
  {
    return arrRegExpItems[ <%=funIntDateFormatPosition( "mm" , objVbsDb )%> ];
  }
  
  ////////////////////////
  function funBolIsGoodDateWithRegExpMinuteOk( arrRegExpItems )
  ////////////////////////
  {
    if ( funBolMinuteIsInserted( arrRegExpItems ) )
    {
      // the user inserted minutes
      var intMinute = funIntRegExpMinute( arrRegExpItems );
      return ( ( intMinute >= 0 ) && ( intMinute <= 59 ) );
    }
    else
      // the user didn't insert minutes
      return true;
  }

  ////////////////////////
  function funBolSecondIsInserted( arrRegExpItems )
  ////////////////////////
  {
    return ( arrRegExpItems.length >= ( <%=funIntDateFormatPosition( "ss" , objVbsDb )%> + 1 ) );
  }

  ////////////////////////
  function funIntRegExpSecond( arrRegExpItems )
  ////////////////////////
  {
    return arrRegExpItems[ <%=funIntDateFormatPosition( "ss" , objVbsDb )%> ];
  }
  
  ////////////////////////
  function funBolIsGoodDateWithRegExpSecondOk( arrRegExpItems )
  ////////////////////////
  {
    if ( funBolSecondIsInserted( arrRegExpItems ) )
    {
      // the user inserted seconds
      var intSecond = funIntRegExpSecond( arrRegExpItems );
      return ( ( intSecond >= 0 ) && ( intSecond <= 59 ) );
    }
    else
      // the user didn't insert seconds
      return true;
  }

  ////////////////////////
  function funBolIsGoodDateWithRegExp( arrRegExpItems )
  ////////////////////////
  {
	return  funBolIsGoodDateWithRegExpYearOk( arrRegExpItems ) &&
	        funBolIsGoodDateWithRegExpMonthOk( arrRegExpItems ) &&
	        funBolIsGoodDateWithRegExpDayOk( arrRegExpItems ) &&
	        funBolIsGoodDateWithRegExpHourOk( arrRegExpItems ) &&
	        funBolIsGoodDateWithRegExpMinuteOk( arrRegExpItems ) &&
	        funBolIsGoodDateWithRegExpSecondOk( arrRegExpItems );
  }

  ////////////////////////
  function funBolIsGoodNonEmptyDate( objField )
  ////////////////////////
  {
	var arrRegExpItems;
	arrRegExpItems = /<%=funStrDateFormatRegularExpression( objVbsDb )%>/.exec( objField.value );
	if ( ( arrRegExpItems == null ) || ( arrRegExpItems.input != arrRegExpItems[ 0 ] ) )
	  // no match
	  return false;
	else
	  // good match
	  return funBolIsGoodDateWithRegExp( arrRegExpItems );
  }
  
  ////////////////////////
  function funBolIsGoodDate( objField )
  ////////////////////////
  {
	return ( ( objField.value == '' ) || funBolIsGoodNonEmptyDate( objField ) )
  }
  
  ////////////////////////
  function vbsDbValidateDate( objField , strMessage )
  ////////////////////////
  {
	var msg_addition = ""
	//new_fieldname = fieldname
	if (!funBolIsGoodDate( objField ) )
	  return strMessage + "\n";
	else
	  return "";
  }
<%
end sub
%>