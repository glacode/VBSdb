<%
'***********************************************
'***********************************************
function funBolInputTemplate( objVbsDb )
'***********************************************
'***********************************************
  funBolInputTemplate = _
    ( VbsDbGetScreenType( objVbsDb ) = "Search" and objVbsDb( "SearchTemplate" ) <> "" ) or _
    ( VbsDbGetScreenType( objVbsDb ) = "Add" and objVbsDb( "EditAddTemplate" ) <> "" ) or _
    ( VbsDbGetScreenType( objVbsDb ) = "Update" and objVbsDb( "EditUpdateTemplate" ) <> "" ) or _
    ( VbsDbGetScreenType( objVbsDb ) = "Delete" and objVbsDb( "EditDeleteTemplate" ) <> "" )
end function

'***********************************************
function funStrInputTemplateFile( objVbsDb )
'***********************************************
  select case VbsDbGetScreenType( objVbsDb )
    case "Search"
      funStrInputTemplateFile = objVbsDb( "SearchTemplate" )
    case "Add"
      funStrInputTemplateFile = objVbsDb( "EditAddTemplate" )
    case "Update"
      funStrInputTemplateFile = objVbsDb( "EditUpdateTemplate" )
    case "Delete"
      funStrInputTemplateFile = objVbsDb( "EditDeleteTemplate" )
    case else
      funStrInputTemplateFile = ""
  end select
end function

'***********************************************
sub vbsDbSetInputTemplateContent( objVbsDb )
'***********************************************
  objVbsDb( "InputTemplateContent" ) = funStrTextFileContent( funStrInputTemplateFile( objVbsDb ) , objVbsDb )
end sub

'***********************************************
'***********************************************
sub vbsDbSetInputTemplateRelatedProperties( objVbsDb )
'***********************************************
'***********************************************
  'if objVbsDb( "InputTemplate" ) <> "" then
  if funBolInputTemplate( objVbsDb ) then
    ' the user wants to display a custom input screen
    vbsDbSetInputTemplateContent objVbsDb
  end if
end sub

'***********************************************
function funStrInputTemplateFormBegin( objVbsDb )
'***********************************************
  funStrInputTemplateFormBegin = funStrInputScreenFormBegin( objVbsDb )
end function

'***********************************************
function funStrInputTemplateTransformationWithoutBadSpacesForFields( strInputTemplateWithoutBadSpaces , objVbsDb )
'***********************************************
  dim strFieldName
  funStrInputTemplateTransformationWithoutBadSpacesForFields = strInputTemplateWithoutBadSpaces
  for each strFieldName in objVbsDb( "DictionaryInputFields" )
  	funStrInputTemplateTransformationWithoutBadSpacesForFields = _
  	  replace( funStrInputTemplateTransformationWithoutBadSpacesForFields , _
  	  "<VbsDbFieldParameter>" & strFieldName & "</VbsDbFieldParameter>" , _
      funStrVbsDbFieldParameterValue( strFieldName , objVbsDb ) , 1 , -1 , 1 )   ' last one is needed to perform a text comparison, so that grid template file field bookmarks become case insensitive
  	funStrInputTemplateTransformationWithoutBadSpacesForFields = _
  	  replace( funStrInputTemplateTransformationWithoutBadSpacesForFields , "<VbsDbFieldValue>" & strFieldName & "</VbsDbFieldValue>" , _
      funStrVbsDbDrawInputElement( strFieldName , objVbsDb ) , 1 , -1 , 1 )   ' last one is needed to perform a text comparison, so that form template file field bookmarks become case insensitive
  next
end function

'***********************************************
function funStrInputTemplateTransformationWithoutBadSpacesForButtons( strInputTemplateWithoutBadSpaces , objVbsDb )
'***********************************************
  funStrInputTemplateTransformationWithoutBadSpacesForButtons = _
    replace( strInputTemplateWithoutBadSpaces , "<VbsDbInputSubmit>" , _
      funStrVbsDbDrawInputSubmit( objVbsDb ) , 1 , -1 , 1 )
  funStrInputTemplateTransformationWithoutBadSpacesForButtons = _
    replace( funStrInputTemplateTransformationWithoutBadSpacesForButtons , "<VbsDbInputReset>" , _
      funStrVbsDbDrawInputReset( objVbsDb ) , 1 , -1 , 1 )
  funStrInputTemplateTransformationWithoutBadSpacesForButtons = _
    replace( funStrInputTemplateTransformationWithoutBadSpacesForButtons , "<VbsDbInputCancel>" , _
      funStrVbsDbDrawInputCancel( objVbsDb ) , 1 , -1 , 1 )
end function

'***********************************************
function funStrInputTemplateTransformationWithoutBadSpaces( strInputTemplateWithoutBadSpaces , objVbsDb )
'***********************************************
  funStrInputTemplateTransformationWithoutBadSpaces = _
    funStrInputTemplateTransformationWithoutBadSpacesForFields( strInputTemplateWithoutBadSpaces , objVbsDb )
  funStrInputTemplateTransformationWithoutBadSpaces = _
    funStrTemplateTransformationWithoutBadSpacesPerformExecutes( funStrInputTemplateTransformationWithoutBadSpaces , objVbsDb )
  funStrInputTemplateTransformationWithoutBadSpaces = _
    funStrInputTemplateTransformationWithoutBadSpacesForButtons( funStrInputTemplateTransformationWithoutBadSpaces , objVbsDb )
end function

'***********************************************
function funStrInputTemplateTransformationBody( objVbsDb )
'***********************************************
  funStrInputTemplateTransformationBody = _
    funStrCutSpacesWithinElement( "VbsDbFieldValue" , objVbsDb( "InputTemplateContent" ) )
  funStrInputTemplateTransformationBody = _
    funStrInputTemplateTransformationWithoutBadSpaces( funStrInputTemplateTransformationBody , objVbsDb )
end function

'***********************************************
function funStrInputTemplateFormEnd( objVbsDb )
'***********************************************
  funStrInputTemplateFormEnd = funStrVbsDbHiddenFields( objVbsDb ) & "</form>"
end function

'***********************************************
function funStrInputTemplateTransformation( objVbsDb )
'***********************************************
  funStrInputTemplateTransformation = _
    funStrInputTemplateFormBegin( objVbsDb ) & _
    funStrInputTemplateTransformationBody( objVbsDb ) & _
    funStrInputTemplateFormEnd( objVbsDb )
end function

'***********************************************
'***********************************************
sub vbsDbDrawTemplateInputScreen( objVbsDb )
'***********************************************
'***********************************************
  response.write funStrInputTemplateTransformation( objVbsDb )
end sub
%>