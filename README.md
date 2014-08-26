# runVBAFilesInOffice


    runVBAFilesInOffice -excel VBS_File_One VBS_File_two ... -c Function argument_one argument_two ...

Run files with *Visual Basic for Application* code/programms in Word, Excel or Visio.

The application in which the code is run is determined by on of the flags `-word`, `-excel` or
`-visio`.

Additionally, the flag `-wsh` can be used to add a reference to the *Windows Script Host Object Model*
(Guid = <code>F935DC20-1CF0-11D0-ADB9-00C04FD58A0B</code>). ([Windows_font-path.bas](https://github.com/ReneNyffenegger/Fonts/blob/master/Windows_font-path.bas)
uses this flag).


## Links

[perl Win32::OLE](https://github.com/ReneNyffenegger/perl-Win32-OLE).

Some examples on [ADODB](https://github.com/ReneNyffenegger/about-adodb/tree/master/Oracle) need *runVBAFilesInOffice*:
[anonymous_block.bas](https://github.com/ReneNyffenegger/about-adodb/blob/master/Oracle/anonymous_block.bas),
[ref_cursor.bas](https://github.com/ReneNyffenegger/about-adodb/blob/master/Oracle/ref_cursor.bas) and
[stored_procedure.bas](https://github.com/ReneNyffenegger/about-adodb/blob/master/Oracle/stored_procedure.bas).
