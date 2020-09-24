<div align="center">

## Set a Printer as default from Printers collection


</div>

### Description

Attempting to set the default printer to an object variable has no effect. For instance, given a system with more than one printer installed, the following code will not change the default printer: Set Printer = Printers(2). The expected behavior is that the document should print to the first non-default printer found in the printers collection. The actual behavior is that the document prints to the original default printer. Thus, a fix was proposed in MSDN Article ID Q167735. I modified this code from it original and wrapped in a class for the purpose of storing the original printer configuration during class initialization and reseting it back during termination if it was modified.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |1999-12-28 22:40:16
**By**             |[Vincent Lozada](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vincent-lozada.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD260212281999\.zip](https://github.com/Planet-Source-Code/vincent-lozada-set-a-printer-as-default-from-printers-collection__1-5187/archive/master.zip)








