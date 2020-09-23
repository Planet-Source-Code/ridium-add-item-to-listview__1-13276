<div align="center">

## Add Item to listview


</div>

### Description

takes a string like "Johnny*(722)555-5555*35" and splits it into columns on a listview control(designed for Report view)
 
### More Info
 
lst = Listview control

zString = data string with *'s between the words


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ridium](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ridium.md)
**Level**          |Beginner
**User Rating**    |4.4 (44 globes from 10 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ridium-add-item-to-listview__1-13276/archive/master.zip)





### Source Code

```
Sub ReportAddTo(lst As ListView, zString As String)
Dim bleh As ListItem
'zString = "One*Two*Three*Four*Five"
On Error Resume Next
Set bleh = lst.ListItems.Add(, , Split(zString, "*")(0))
For i = 1 To 200
  bleh.SubItems(i) = Split(zString, "*")(i)
Next i
End Sub
```

