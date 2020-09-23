<div align="center">

## Store Anything in a Tag


</div>

### Description

People often think they can only store one item in a tag and that that item must be a string. Not so. A tag is a string limited in length only by memory available - we can treat it as an array of bytes and store any data in it including multiple items. The best way to illustrate this is by an example.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kenneth Buckmaster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kenneth-buckmaster.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kenneth-buckmaster-store-anything-in-a-tag__1-68205/archive/master.zip)

### API Declarations

Copymemory


### Source Code

```

'in a form
'set autoredraw to true
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'in our example we'll store
'a date + a long
'+ a string
'(as per a normal tag)
'at the end
'the date is 8 bytes,
'the long 4 so we need
'the tag to be 12 bytes
'= 6 characters
'to be able store these
Sub initialiseTag(obj As Object)
obj.Tag = String(6, Chr(0)) & obj.Tag
End Sub
'the first item in the tag
'is going to be a date
Sub setTagDate(obj As Object, dte As Date)
'we can't copymemory directly to the tag
'we must use a temporary string
Dim tmpstr As String
tmpstr = obj.Tag
'make sure the tag has been initialised
If Len(tmpstr) > 5 Then
'copy date to tmpstr
CopyMemory ByVal StrPtr(tmpstr), ByVal VarPtr(dte), 8 'date = 8 bytes
End If
'set tag to altered tmpstr
obj.Tag = tmpstr
End Sub
Function getTagDate(obj As Object) As Date
If Len(obj.Tag) > 5 Then
'we can use copymemory to copy from the tag
CopyMemory ByVal VarPtr(getTagDate), ByVal StrPtr(obj.Tag), 8
End If
End Function
'same for the second item we'll store - a long
Sub setTagLong(obj As Object, ll As Long)
Dim tmpstr As String
tmpstr = obj.Tag
If Len(tmpstr) > 5 Then
'only this time we need to add
'the sum of the previous items
'to strptr
'previous was one date = 8 bytes
CopyMemory ByVal StrPtr(tmpstr) + 8, ByVal VarPtr(ll), 4 'long = 4 bytes
End If
obj.Tag = tmpstr
End Sub
Function getTagLong(obj As Object) As Long
If Len(obj.Tag) > 5 Then
CopyMemory ByVal VarPtr(getTagLong), ByVal StrPtr(obj.Tag) + 8, 4
End If
End Function
'last item well use like a normal tag
Sub setTagString(obj As Object, tagstring As String)
If Len(obj.Tag) > 5 Then
'6 character = 12 bytes = length of the date and long we're storing
'before this string
obj.Tag = Left(obj.Tag, 6) & tagstring
End If
End Sub
Function getTagString(obj As Object) As String
If Len(obj.Tag) > 6 Then
getTagString = Right(obj.Tag, Len(obj.Tag) - 6)
End If
End Function
Private Sub Form_Load()
'allocate space for the fixed length data
initialiseTag Me
'set values
setTagString Me, "String"
setTagDate Me, Date
setTagLong Me, -87454
'recall Values
Me.Print getTagDate(Me)
Me.Print getTagLong(Me)
Me.Print getTagString(Me)
End Sub
```

