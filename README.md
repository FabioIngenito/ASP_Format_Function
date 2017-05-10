# ASP_Format_Function
CRIANDO UM FORMAT MASCARA PARA O ASP

'Este código "cria" um "format" para o ASP! Você pode usar o que quiser substituíndo o caracter "#" pelo valor que precisa formatar.
'Thank you, Mr.Brian Reeves!

'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=8175&lngWId=4

'Can't Copy and Paste this?
'Click here for a copy-and-paste friendly version of this code!
'**************************************
' for :ASP Format Function
'**************************************
'Open Source

'Terms of Agreement:   
'By using this code, you agree to the following terms...   
'1. You may use this code in your own programs (and may compile it into a program and distribute it in compiled format for languages that allow it) freely and with no charge. 
'2. You MAY NOT redistribute this code (for example to a web site) without written permission from the original author. Failure to do so is a violation of copyright laws.    
'3. You may link to this code from another website, but ONLY if it is not wrapped in a frame.  
'4. You will abide by any additional copyright restrictions which the author may have placed in the code or code's description. 
                               
'**************************************
' Name: ASP Format Function
' Description:This function operates similarly to the VB Format function with one big exception. The "#" character is used to represent any single character. You can trim all non alphanumeric characters out and reformat them to stay consistant.
Usefull for credit cards, zipcodes, phone numbers, etc...
' By: Brian Reeves
'
' Assumes:Format("1234567890123", "(###) ###-#### x######") would return "(123) 456-7890 x123"
Format("4111111111111111", "####-####-####-####")
would return "4111-1111-1111-1111"
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=8175&lngWId=4'for details.'**************************************

'******
'**            Formats a string to include standard sets.
'**
'**            Example:       Format("1234567890", "(###) ###-####")
'**                    Result =       (123) 456-7890
'**            Modified 01/09/03 to allow extended format mask that will
'**                    not return extra ###'s brian reeves
'******

Public Function Format(sValue, sMask)
        Dim iPlaceHolder
        Dim sTempValue
        Dim sResult
        
        sTempValue = CStr(sValue)
        sResult = sMask

        Do Until InStr(sResult, "#") = 0
               iPlaceHolder = InStr(sResult, "#")
               sResult = Replace(sResult, "#", Left(sTempValue, 1), 1, 1)
               sTempValue = Mid(sTempValue, 2)
               If Len(sTempValue) = 0 Then sResult = Left(sResult, iPlaceHolder)
        Loop

        Format = sResult
End Function        
