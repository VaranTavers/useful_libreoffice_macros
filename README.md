# Varan Tavers' useful macros for LibreOffice

This is a selection of LibreOffice macros that I find useful for my work.

Since very few people actually like writing BASIC (and I'm not one of them) this code was written assisted by ChatGPT.

## Selection to LaTeX table (CalcToLatex.bas)

Creates a LaTeX table from the selection, and inserts the code inside a new text-box.

Maybe unexpected things it does:

* Transforms all commas into periods in numbers.
* Escapes underscore characters.

When the EasyMacro extension worked, it copied it directly into the clipboard, that functionality is not working right now.


## Export all sheets as PDF-s (ExportAll)

Exports all sheets into separate PDF-s into the same directory.