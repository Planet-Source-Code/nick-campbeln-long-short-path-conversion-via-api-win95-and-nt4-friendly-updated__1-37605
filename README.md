<div align="center">

## Long \<\-\> Short Path Conversion via API \- Win95 and NT4 Friendly \(Updated\)


</div>

### Description

This code module gives you the ability to convert between short (DOS style 8.3) and long paths under ANY Win32 system. If you've ever had to do path conversions with the API you know that Win95/NT4 does not support the GetLongPathName() API which allows for short to long path name conversions. Those of you who have looked for a solution may have found David Goben’s wonderful (though poorly commented) GetLongPath() function on SearchVB.com, but it relies on the FileSystemObject (which did me no good). So I built this module which utilizes both conversion APIs (GetShortPathName() and GetLongPathName()) on the systems that support them (Win98+). For systems that do not support GetLongPathName() (Win95/NT4) the GetShortPathName() API is used along with Dir() to determine the long path name from the passed short path name. The code is very well commented but is not thoroughly tested; so if you find a bug, please let me know!. Thanks and enjoy!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-08-26 12:32:24
**By**             |[Nick Campbeln](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nick-campbeln.md)
**Level**          |Beginner
**User Rating**    |5.0 (45 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Long\_\_\-\_\_S1227648262002\.zip](https://github.com/Planet-Source-Code/nick-campbeln-long-short-path-conversion-via-api-win95-and-nt4-friendly-updated__1-37605/archive/master.zip)








