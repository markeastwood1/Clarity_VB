# Clarity Form Repository

This project contains the VB code used in the Clarity form. The files are named for the pages to which they relate.

The file named `worksheet.vb` is generic and applies to all sheets that don't have a file associated to them.

The file named `thisworkbook.vb` is is used by the workbook level. In the Excel VB editor you can use the project explorer to navigate to this. I try to put very common things up there to avoid repeating them on individual pages.

## Protection
The code itself is locked using the default password as are cells on all sheets except where we've specifically unlocked them for user input. The idea is simple, don't allow people to mess with the form. Also, "structure protection" at teh workbook level means they can't hide/un-hide tabs either, its totally controlled by the VB code.

Eduardo was here 
