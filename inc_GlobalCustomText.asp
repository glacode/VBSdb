<%
'***********************************************
'***********************************************
function funStrGetCustomText( strTextItemName , objVbsDb )
'***********************************************
'***********************************************
  funStrGetCustomText = objVbsDb( "DictionaryGlobalCustomText" )( uCase( strTextItemName ) )
end function

'***********************************************
'***********************************************
function funBolCustomTextExists( strTextItemName , objVbsDb )
'***********************************************
'***********************************************
  funBolCustomTextExists = objVbsDb( "DictionaryGlobalCustomText" ).exists( uCase( strTextItemName ) )
end function

'***********************************************
'***********************************************
sub vbsDbSetDictionaryGlobalCustomText( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "GlobalCustomText" , objVbsDb
end sub
%>