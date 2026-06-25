<%
'***********************************************
sub vbsDbSetDictionaryLanguages( objVbsDb )
'***********************************************
  set objVbsDb( "DictionaryLanguages" ) = server.createObject( "scripting.dictionary" )
  objVbsDb( "DictionaryLanguages" ).compareMode = 1
%>
  <!--#include file="languages/inc.asp"-->
<%
end sub

'***********************************************
function funStrItemLanguage( strItemName )
'***********************************************
  dim intUnderscorePosition
  intUnderscorePosition = inStr( strItemName , "_" )
  funStrItemLanguage = left( strItemName , intUnderscorePosition - 1 )
end function

'***********************************************
function funBolIsChosenLanguage( strItemName , objVbsDb )
'***********************************************
  funBolIsChosenLanguage = ( uCase( funStrItemLanguage( strItemName ) ) = _
                             uCase( objVbsDb( "GlobalLanguage" ) ) )
end function

'***********************************************
function funStrLanguageItemName( strItemName )
'***********************************************
  dim intUnderscorePosition
  intUnderscorePosition = inStr( strItemName , "_" )
  funStrLanguageItemName = right( strItemName , len( strItemName ) - intUnderscorePosition )
end function

'***********************************************
sub vbsDbSetDictionaryLanguageAddItem( strItemName , objVbsDb )
'***********************************************
  if funBolIsChosenLanguage( strItemName , objVbsDb ) then
    ' current item is of the chosen language
    objVbsDb( "DictionaryLanguage" ).add uCase( funStrLanguageItemName( strItemName ) ) , _
      objVbsDb( "DictionaryLanguages" )( strItemName )
  end if
end sub

'***********************************************
'***********************************************
sub vbsDbSetDictionaryLanguage( objVbsDb )
'***********************************************
'***********************************************
  dim strItemName , strTranslation
  vbsDbSetDictionaryLanguages objVbsDb
  set objVbsDb( "DictionaryLanguage" ) = server.createObject( "scripting.dictionary" )
  objVbsDb( "DictionaryLanguage" ).compareMode = 1
  for each strItemName in objVbsDb( "DictionaryLanguages" )
    vbsDbSetDictionaryLanguageAddItem strItemName , objVbsDb
  next
end sub

'***********************************************
'***********************************************
function funStrTranslate( strTextItemName , objVbsDb )
'***********************************************
'***********************************************
  if funBolCustomTextExists( strTextItemName , objVbsDb ) then
    ' user specified a custom text value for this text item
    funStrTranslate = objVbsDb( "DictionaryGlobalCustomText" )( uCase( strTextItemName ) )
  elseif objVbsDb( "DictionaryLanguage" ).exists( uCase( strTextItemName ) ) then
    ' user didn't specify a custom text value for this text item
    funStrTranslate = objVbsDb( "DictionaryLanguage" )( uCase( strTextItemName ) )
  else
    ' user didn't specify a custom text value for this text item and no traslation was provided for the chosen language
    drawError "missing translation for item " & strTextItemName & "." , objVbsDb
  end if
end function
%>