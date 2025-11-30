# Manual_translator
This is a "transaltor" that you make yourself - it takes each word from the form input and tries to find it in an Excel file where this macro rests, and then returns the translated version. Every word needs to be added manually and there's no spellcheck, which is sadly the best way of doing the task for which this macro was made. 
There's also a couple of QOL macros, like replacing german-specific letters, cleaning insignificant spaces or changing how some specific numbers are written.

## Macros
### translator
Open the Excel in which the macro and your dictionary are, and then run the "translator" macro. A form window will pop up, in which you just need to paste the string that you want to translate. After confirming input, it will show the answer below and copy it to the clipboard. In case that you need to translate not from left column to right column, but vice versa, just put a "-" minus sign as the first symbol in input.
### workbookCleaner 
Open an Excel document and activate this macro. On every sheet in the workbook it will: <br />
&emsp;replace all german-specific letters so it won't cause problems <br />
&emsp;delete all strikedthrough text, all types of brackets, stars *, underscores _ and non-breaking-spaces <br />
&emsp;make all text Uppercase <br />
&emsp;delete leading and trailing spaces <br />
&emsp;remove extra spaces
### AXXXXXXXXXX or A_XXX_XXX_XX_XX
It does exactly what it promises - for each cell that has english letter A followed by 10 digits, it either insert spaces as seen in macro's name, or deletes all spaces.


## Chnage Log
V1.0 - initial release
