<%
'***********************************************
function funStrFormFieldNameWithoutBlanks( strFieldName , objVbsDb )
'***********************************************
  funStrFormFieldNameWithoutBlanks = replace( strFieldName , " " , constStrBlankReplacementsForInputFormControlNames )
end function

'***********************************************
'***********************************************
function funStrFormFieldName( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  dim strFieldNameWithoutBlanks
  strFieldNameWithoutBlanks = funStrFormFieldNameWithoutBlanks( strFieldName , objVbsDb )
  'strFieldNameWithoutBlanks = replace( strFieldNameWithoutBlanks , "." , constStrDotReplacementForInputFormControlNames )
  if VbsDbGetScreenType( objVbsDb ) = "Search" then
    funStrFormFieldName = "vbsDbFilterField_" & strFieldNameWithoutBlanks
  else
    ' the screen is an add or an update
    funStrFormFieldName = strFieldNameWithoutBlanks
  end if
end function

'***********************************************
'***********************************************
function funStrVbsDbHiddenFieldTypeName( strInputFieldName , objVbsDb )
'***********************************************
'***********************************************
  funStrVbsDbHiddenFieldTypeName = "VBSdbFieldType_vbsDbFilterField_" & funStrFormFieldNameWithoutBlanks( strInputFieldName , objVbsDb )
end function

'***********************************************
'***********************************************
function funBolVbsDbIsFilterField( strFormControlName , objVbsDb )
'***********************************************
'***********************************************
  funBolVbsDbIsFilterField =  left( strFormControlName , 17 ) = "vbsDbFilterField_"
end function

'***********************************************
'***********************************************
function funStrFilterFieldName( strFormControlName )
'***********************************************
'***********************************************
  funStrFilterFieldName = mid( strFormControlName , 18 )
  funStrFilterFieldName = replace( funStrFilterFieldName , constStrBlankReplacementsForInputFormControlNames , " " )
  'funStrFilterFieldName = replace( funStrFilterFieldName , constStrDotReplacementForInputFormControlNames , "." )
end function
%>