<%
'***********************************************
'***********************************************
sub vbsDbSetDictionaryObjectForViewFieldFormats( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "ViewFieldFormats" , objVbsDb
end sub

'***********************************************
function funStrVbsDbCustomViewFieldFormat( strFieldName , objVbsDb )
'***********************************************
	dim objField
	funStrVbsDbCustomViewFieldFormat = objVbsDb( "DictionaryViewFieldFormats" )( uCase( strFieldName ) )
	for each objField in objVbsDb( "SqlRecordSet" ).fields
  	funStrVbsDbCustomViewFieldFormat = replace( funStrVbsDbCustomViewFieldFormat , _
  	  "<VbsDbFieldParameter>" & objField.name & "</VbsDbFieldParameter>" , _
      funStrVbsDbFieldParameterValue( objField.name , objVbsDb ) , 1 , -1 , 1 )   ' last one is needed to perform a text comparison, so that grid template file field bookmarks become case insensitive
	  'funStrVbsDbCustomViewFieldFormat = replace( funStrVbsDbCustomViewFieldFormat , """""{{" & objField.name & "}}""""" , _
	  '  replace( funStrConvertNullToEmptyString( objField.value ) , """" , """" & """" , 1 , -1 , 1 ) , 1 , -1 , 1 )   ' to handle text parameters for the server side Execute
	  funStrVbsDbCustomViewFieldFormat = replace( funStrVbsDbCustomViewFieldFormat , "{{" & objField.name & "}}" , _
	    funStrConvertNullToEmptyString( objField.value ) , 1 , -1 , 1 )   ' last one is needed to perform a text comparison, so that ViewFieldFormats property becomes case insensitive
	next
end function

'***********************************************
function funStrHtmlForCarriageReturn( strMemoFieldValue )
'***********************************************
  funStrHtmlForCarriageReturn = replace( strMemoFieldValue , chr( 10 ) , "<br>" )
end function

'***********************************************
function funStrVbsDbViewFieldFormatDontCustomize( strFieldValue )
'***********************************************
  funStrVbsDbViewFieldFormatDontCustomize = server.htmlEncode( strFieldValue )
  funStrVbsDbViewFieldFormatDontCustomize = funStrHtmlForCarriageReturn( funStrVbsDbViewFieldFormatDontCustomize )
end function

'***********************************************
'***********************************************
function funStrVbsDbViewFieldFormat( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "DictionaryViewFieldFormats" )( uCase( strFieldName ) ) <> "" then
  	' user wants to customize this field format
  	funStrVbsDbViewFieldFormat = funStrVbsDbCustomViewFieldFormat( strFieldName , objVbsDb )
  	funStrVbsDbViewFieldFormat = funStrTemplateTransformationWithoutBadSpacesPerformExecutes( funStrVbsDbViewFieldFormat , objVbsDb )
  else
  	' user doesn't want to customize this field format
  	funStrVbsDbViewFieldFormat = funStrVbsDbViewFieldFormatDontCustomize( funStrConvertNullToEmptyString( objVbsDb( "SqlRecordSet" )( strFieldName ).value ) )
  end if
end function
%>