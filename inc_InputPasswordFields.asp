<%
'***********************************************
'***********************************************
function funBolFieldIsOfPasswordType( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  funBolFieldIsOfPasswordType = objVbsDb( "DictionaryInputPasswordFields" ).exists( strFieldName )
end function

'***********************************************
'***********************************************
sub vbsDbSetDictionaryInputPasswordFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "InputPasswordFields" , objVbsDb
end sub
%>