Attribute VB_Name = "modremover"
'Use this module to remove HTML from a byte array
'It's speed is very high, when compiled it achives
'an average 1 MB/sec on a Pentium I 500 mhz

Const START_TAG_CODE = 60 'Asc("<")
Const END_TAG_CODE = 62 'Asc(">")
Const START_SYMBOL_CODE = 38 'Asc("&")
Const END_SYMBOL_CODE = 59 'Asc(";")

'Allows passing control to events
Dim allowevents As Boolean

Function RemoveHTML(ByteArray() As Byte)
Dim tagfounded As Boolean
tagfounded = False
Dim OutputText() As Byte
ReDim OutputText(500)
Dim Bytes As Long
Bytes = 0
For cnt = 0 To UBound(ByteArray)

'Check if there is no more space to store the text
If Bytes = UBound(OutputText) Then
'Reserve 500 bytes more of space
ReDim Preserve OutputText(Bytes + 500)
End If

'If an HTML tag was founded
If tagfounded = True Then
'Check if we founded the end of the tag
If ByteArray(cnt) = END_TAG_CODE Or ByteArray(cnt) = END_SYMBOL_CODE Then
'We founded the end of the tag
tagfounded = False
End If
'Check if current byte is the start of an HTML tag
ElseIf ByteArray(cnt) = START_TAG_CODE Or ByteArray(cnt) = START_SYMBOL_CODE Then
'Set that the an HTML tag was founded
tagfounded = True
'If allowing passing control to events
If allowevents = True Then
'Pass control to events when we found an HTML tag
DoEvents
End If
End If

'Check if we didn't found any HTML tag
If tagfounded = False Then
If ByteArray(cnt) <> END_TAG_CODE And ByteArray(cnt) <> END_SYMBOL_CODE Then
'Output current byte
OutputText(Bytes) = ByteArray(cnt)
'Increment the number of written bytes
Bytes = Bytes + 1
End If
End If
Next

'Redim the output text to the correct size
ReDim Preserve OutputText(Bytes - 1)

'Return the output text
RemoveHTML = OutputText()
End Function
