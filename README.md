<div align="center">

## UUDecode Class


</div>

### Description

This code is based upon the ONLY other VB uuencode/uudecode source code available (it's here). I put this together as part of a larger project that I'm in the middle of (a USENET binary downloader) and I wanted a uudecode routine that would decode on-the-fly. This does. It also encodes and decodes files, like a normal encoder/decoder. I spent a lot of time optimizing it, and I think it'll now decode fast enough to handle most DSL/Cable accounts on the fly (480-590 kb/s). Error checking is sparse, hopefully there's no glaring bugs. Hope it's helpful, and let me know if there's any more optimizations that you can find!

[Update] I fixed the problems described below, and also replaced the static filenumbers with dynamically generated ones. Doh!

[Update] Another bug found and fixed. There were null characters at the end of each line of files encoded with UUEncodeFile. They're gone now...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-04-05 08:07:40
**By**             |[Mark Parker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-parker.md)
**Level**          |Advanced
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD4545452000\.zip](https://github.com/Planet-Source-Code/mark-parker-uudecode-class__1-6879/archive/master.zip)








