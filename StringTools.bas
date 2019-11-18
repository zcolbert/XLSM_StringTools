Attribute VB_Name = "StringTools"
' Module Name:  StringTools
' Purpose:      A collection of string manipulation functions
' Author:       Zac Colbert
' Contact:      zcolbert1993@gmail.com
'
' Modified:     07/19/2019 - Fixed name conflict in REPLACESTR
'               08/14/2019 - Added optional delimiter to WORD function
'               11/18/2019 - Change name of index parameter in WORD function


Function AFTER(text As String, delimiter As String, Optional startIndex As Variant)
    ' Return text occurring after the Nth occurrence of the delimiter
    '   text:       The string to be searched
    '   delimiter:  The delimiting string
    '   startIndex: The starting index of the delimiter (default=1)
    
    ' Set default starting index if parameter not supplied
    If IsMissing(startIndex) Then
        startIndex = 1
    End If

    ' Split text string at delimiter
    Dim words As Variant
    words = Split(text, delimiter)
    
    ' Define array to hold the words after the Nth delimiter
    Dim result() As String
    ReDim result(UBound(words) - startIndex)  ' Resize results array

    ' Load the results array with words after Nth delimiter
    Dim currentIndex As Integer
    resultIndex = 0
    For i = startIndex To UBound(words)
        result(resultIndex) = words(i)
        resultIndex = resultIndex + 1
    Next i
    
    AFTER = Join(result, delimiter)

End Function


Function BEFORE(text As String, delimiter As String, Optional occurrence As Variant)
    ' Return text before Nth occurrence of delimiter
    '   text:       The text to be searched
    '   delimiter:  The delimiting string
    '   occurrence: Index of delimiter used as a stopping point
    
    If IsMissing(occurrence) Then
        occurrence = 1
    End If
    
    Dim words As Variant
    words = Split(text, delimiter)
    
    Dim lastIndex As Integer
    lastIndex = occurrence - 1

    ' Resize array to the proper number of elements
    ' preserving existing contents up to lastIndex
    ReDim Preserve words(lastIndex)

    ' Join words into a new string
    BEFORE = Join(words, delimiter)
    
End Function


Function BETWEEN(text As String, leftDelim As String, rightDelim As String)
    ' Return the string located between leftDelim and rightDelim
    '   text:       The source string
    '   leftDelim:  The string delimiter indicating the start position
    '   rightDelim: The string delimiter indicating the end position
    
    If leftDelim = rightDelim Then
        ' Delimiters are identical
        ' Return text between adjacent occurrences of delimiter
        Dim words As Variant
        words = Split(text, leftDelim)
        BETWEEN = words(1)
    Else
        ' Delimiters are unique
        ' Return text between startDelim and endDelim
        Dim startIndex As Integer
        Dim endIndex As Integer
        
        startIndex = InStr(text, leftDelim) + Len(leftDelim)
        endIndex = InStr(text, rightDelim)
        
        Dim length As Integer
        length = endIndex - startIndex
        
        BETWEEN = Mid(text, startIndex, length)
    End If

End Function


Function FIRST(text As String, Optional count As Variant)
    ' Return the first words(s) from a space delimited string
    '   text:   The source text
    '   count:  The number of words to return (default=1)
    '           counting from the beginning of the string.
    '           If count >= the total number of words,
    '           the entire string is returned.
    
    If IsMissing(count) Then
        count = 1
    End If

    FIRST = BEFORE(text, " ", count)
    
End Function


Function LAST(text As String, Optional count As Variant)
    ' Return the last word(s) from a space delimited string
    '   text:   The source text
    '   count:  The number of words to return (default=1)
    '           counting backwards from the end of the string.
    '           If count >= the total number of words,
    '           the entire string is returned.
    
    If IsMissing(count) Then
        count = 1
    End If
        
            
    Dim delimiter As String
    delimiter = " "
        
    If count = 1 Then
        ' Return text after right-most occurrence of the delimiter
        
        ' Find last occurrence of delimiting character
        Dim lastDelimPos As Integer
        lastDelimPos = InStrRev(text, delimiter)
        
        ' Return text occurring after last delimiting character
        LAST = Right(text, Len(text) - lastDelimPos)
        
    Else
        ' Split the text at each occurrence of the delimiter
        Dim words As Variant
        words = Split(text, delimiter)

        ' Create an array to hold the last (count) words
        Dim result() As String
        ReDim result(count)  ' Resize result

        ' Load results array from source text
        Dim startIndex As Integer
        startIndex = UBound(words) - (count - 1)

        Dim resultIndex As Integer
        resultIndex = 0
        
        For i = startIndex To UBound(words)
            result(resultIndex) = words(i)
            resultIndex = resultIndex + 1
        Next i
        
        LAST = Join(result, delimiter)
    
    End If

End Function


Function ISPLIT(text As String, delimiter As String, index As Integer)
    ' Split a string and return the element located at {index}
    '   text:       The string to split
    '   delimiter:  Text is split at each occurrence of this substring
    '   index:      The index of the element to be returned, starting at 1
    
    index = index - 1  ' Adjust for zero indexed function
    ISPLIT = Split(text, delimiter)(index)
    
End Function


Function LSTRIP(text As String, substr As String)
    ' Remove substring from start of text
    '   text:   The source text
    '   substr: The substring to be removed
    
    If Left(text, Len(substr)) = substr Then
        text = Right(text, Len(text) - Len(substr))
    End If
    
    LSTRIP = text
    
End Function


Function RSTRIP(text As String, substr As String)
    ' Remove substring from end of text
    '   text:   The source text
    '   substr: The substring to be removed
    
    If Right(text, Len(substr)) = substr Then
        text = Left(text, Len(text) - Len(substr))
    End If
    
    RSTRIP = text

End Function


Function STRIP(text As String, substr As String)
    ' Remove substr from both ends of text
    '   text:   The source text
    '   substr: The substring to be removed
    
    text = LSTRIP(text, substr)
    text = RSTRIP(text, substr)
           
    ' Return stripped text
    STRIP = text

End Function


Function REMOVESTR(text As String, removeText As String)
    ' Remove all occurrences of {removeText} from the source {text}
    '   text:        The source text
    '   removeText:  The string to be removed
    
    ' Replace removeStr with empty string
    REMOVESTR = REPLACESTR(text, removeText, "")

End Function


Function REPLACESTR(text As String, oldStr As String, newStr As String)
    ' Replace all occurrences of {oldStr} in {text} with {newStr}
    '   text:       The source text
    '   oldStr:     The string to be replaced
    '   newStr:     The string to be inserted

    REPLACESTR = Join(Split(text, oldStr), newStr)

End Function


Function WORD(text As String, index As Integer, Optional delimiter As Variant)
    ' Return the Nth word in a string
    '   text:       The string to be searched
    '   index:      The index of the word to be returned
    '               If index is negative, count from the
    '               end of the string.
    '   delimiter:  The character separating each word
    '               default is " " (single space)
    
    If IsMissing(delimiter) Then
        delimiter = " "
    End If
    
    Dim words As Variant
    words = Split(text, delimiter)
    
    Dim wordIndex As Integer
    If index < 0 Then
        ' Count backwards from the right-side of word array
        Dim lastIndex As Integer
        index = index * -1
        lastIndex = UBound(words)
        wordIndex = lastIndex - index + 1
    Else
        ' Count from beginning of word array
        wordIndex = index - 1  ' Adjust for zero-index array
    End If
    
    WORD = words(wordIndex)
 
End Function
