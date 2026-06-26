<%
'***********************************************
'***********************************************
function funStrVbsDbSearchOperator( strFieldName , byRef objVbsDb )
'***********************************************
'***********************************************
  funStrVbsDbSearchOperator = "="
  if objVbsDb( "SearchOperators" ) <> "" then
    if objVbsDb( "DictionarySearchOperators" )( strFieldName ) <> "" then
      if objVbsDb( "DictionarySearchOperators" )( strFieldName ) <> "=" and _
         objVbsDb( "DictionarySearchOperators" )( strFieldName ) <> "<" and _
         objVbsDb( "DictionarySearchOperators" )( strFieldName ) <> "<=" and _
         objVbsDb( "DictionarySearchOperators" )( strFieldName ) <> ">" and _
         objVbsDb( "DictionarySearchOperators" )( strFieldName ) <> ">=" then
        drawError "SearchOperators property bad assignment.<br><br>The field " & strFieldName & _
              " has not been associated to an accepted search operator.<br><br> " &_
              "Accepted values are < , <= , > , >=", objVbsDb
      else
        funStrVbsDbSearchOperator = objVbsDb( "DictionarySearchOperators" )( strFieldName )
      end if
    end if
  end if
end function

'***********************************************
'***********************************************
sub vbsDbNormalizeSeparatorsForSearchOperators( byRef objVbsDb )
'***********************************************
'***********************************************
  objVbsDb( "SearchOperators" ) = replace( objVbsDb( "SearchOperators" ) , "," , ";" )
end sub

'***********************************************
'***********************************************
sub vbsDbSetDictionarySearchOperators( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "SearchOperators" , objVbsDb
end sub
%>