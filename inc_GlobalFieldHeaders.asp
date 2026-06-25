<%
'***********************************************
'***********************************************
sub vbsDbSetDictionaryGlobalFieldHeaders( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "GlobalFieldHeaders" , objVbsDb
end sub

'***********************************************
'***********************************************
function funStrFieldHeader( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "DictionaryGlobalFieldHeaders" )( uCase( strFieldName ) ) = "" then
    ' no name map was defined for this field
    funStrFieldHeader = strFieldName
  else
    ' objVbsDb( "DictionaryGlobalFieldHeaders" )( strFieldName ) is the name map for this field
    funStrFieldHeader = objVbsDb( "DictionaryGlobalFieldHeaders" )( uCase( strFieldName ) )
  end if
end function
%>