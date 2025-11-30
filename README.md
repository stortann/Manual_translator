# Manual_translator
This is a "transaltor" that you make yourself - it takes each word from the form input and tries to find it in an Excel file where this macro rests, and then returns the translated version. Every word needs to be added manually and there's no spellcheck, which is sadly the best way of doing the task for which this macro was made. 
There's also a couple of QOL macros, like replacing german-specific letters, cleaning insignificant spaces or changing how some specific numbers are written.

## Macros
### translator
After running this macro, a form window, in which you just need to paste the string that you want to translate, will pop up. After confirming input, it will show the answer below and copy it to the clipboard. In case that you need to translate not from left column to right column, but vice versa, just put a "-" minus sign as the first symbol in input. It will take those words from the Dictionary sheet. If no translation is found, it won't change the word.
For example:
In Dictionary sheet we have GERMAN words in A column, and ENGLISH words in column B. If input will be BROT, the output will be BREAD.

GERMAN | ENGLISH 
--- | ---
AUTO | CAR
BROT | BREAD
TAUBE | PIGEON

### updateDictionary
When activated, it will create new sheet Dictionary if doesn't exist yet. If it exists, it will just add words to it. It will go through first row of every worksheet searching for cells with text "GERMAN". If found, it will start looking at cells under them. For every "word" - symbol/s without a space between, if not found in a first column of Dictionary sheet, it will copy the "translation" - a cell to the right of a cell in which the "word" resides. It will then append "word" to the first column and copied "translation" to the second column of a Dictionary sheet, both of them will be highlighted red. The result may look like a table above. 

### workbookCleaner 
Open an Excel document and activate this macro. In the active sheet in the workbook it will: <br />
&emsp;replace all german-specific letters so it won't cause problems <br />
&emsp;delete all strikedthrough text, all types of brackets, stars *, underscores _ and non-breaking-spaces <br />
&emsp;make all text Uppercase <br />
&emsp;delete leading and trailing spaces <br />
&emsp;remove extra spaces

### AXXXXXXXXXX or A_XXX_XXX_XX_XX
It does exactly what it promises - for each cell in the active sheet that has english letter A followed by 10 digits, and/or spaces and/or stars * in different places, it will either insert spaces as seen in macro's name, or deletes all spaces.

## Change Log
### V2.0 - added updateDictionary(), changed scope of other macros to be limited to only 1 sheet instead of a whole workbook, small changes to almost everything in MyModule
### V1.0 - initial release
