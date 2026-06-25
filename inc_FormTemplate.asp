<%
'***********************************************
sub vbsDbSetFormTemplateContent( objVbsDb )
'***********************************************
  objVbsDb( "FormTemplateContent" ) = funStrTextFileContent( objVbsDb( "FormTemplate" ) , objVbsDb )
end sub

'***********************************************
'***********************************************
sub vbsDbSetFormTemplateRelatedProperties( objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "FormTemplate" ) <> "" then
    ' the user wants to dispalay a custom grid
    vbsDbSetFormTemplateContent objVbsDb
  end if
end sub

'***********************************************
'function funStrFormTemplateTransformationWithoutBadSpaces( strFormTemplateWithoutBadSpaces , objVbsDb )
'***********************************************
'  dim strFieldName
'  funStrFormTemplateTransformationWithoutBadSpaces = strFormTemplateWithoutBadSpaces
'  for each strFieldName in objVbsDb( "DictionaryFormFields" )
'  	funStrFormTemplateTransformationWithoutBadSpaces = replace( funStrFormTemplateTransformationWithoutBadSpaces , "<VbsDbFieldValue>" & strFieldName & "</VbsDbFieldValue>" , _
'    funStrFieldValueForViewScreen( strFieldName , objVbsDb ) , 1 , -1 , 1 )   ' last one is needed to perform a text comparison, so that form template file field bookmarks become case insensitive
'  next
'end function

'***********************************************
'function funStrFormTemplateTransformation( objVbsDb )
'***********************************************
'  funStrFormTemplateTransformation = funStrCutSpacesWithinElement( "VbsDbFieldValue" , objVbsDb( "FormTemplateContent" ) )
'  funStrFormTemplateTransformation = funStrFormTemplateTransformationWithoutBadSpaces( funStrFormTemplateTransformation , objVbsDb )
'end function

'***********************************************
'***********************************************
sub vbsDbDrawTemplateForm( objVbsDb )
'***********************************************
'***********************************************
  'response.write funStrTemplateGridRowTransformation( objVbsDb )
  response.write funStrVbsDbTemplateTransformation( "Form" , objVbsDb )
end sub
%>