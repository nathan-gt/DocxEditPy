# DocxEditPy

This small python program will take a .docx file and modify its text based on the dictionary provided, then save the new file in the working directory.

An executable can be found in the ```src``` folder, and will take ```src/dictionary.json``` for its input, to test the progam you can use ```src/demo-template.docx``` as the source file

## If strings aren't being replaced
While playing around with ```demo-template.docx``` to see what can and can't be affected by this program, I've noticed a couple of oddballs where strings weren't getting updated. But when rewriting them, instead of copy pasting them like I was previously doing, it worked more. The problem probably comes from the fact a replaced string cannot be of different styles, this includes the tiniest of details like line height.
- **Solution**: Make sure to remove all styling of the string you want to replace in the .docx file with the style eraser tool and then restyle it like you want, make sure every character is the same style and be careful of whitespaces.

## Notes about the executable
For some reason the .exe takes a bit to load, so be patient if nothing appears on the console afer a couple of seconds
 
## Compatibility
Currently only works on windows because of the use of the os library to read and open files but another cross plateform library could be used

### Libraries used:
- python-docx
- pyinstaller (for the .exe)
