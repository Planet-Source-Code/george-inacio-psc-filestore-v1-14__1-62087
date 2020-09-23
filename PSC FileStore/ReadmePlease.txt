
PSC FileStore is design to work only with pages and files
from Planet Source Code.
But you can change it to suite your needs.

PSC FileStore needs a resizer but it’s not included
so you can use your favourite resizer.

On this code, the comments are next to nothing.
This code is meant to be a tool and not a tutorial.
My apologies!

Acknowledgments.
I used code or tools from (In alphabetic order);
AllAPI Guide, Info-Zip, Jim Jose, Lee Weiner, Raymond L. King, Roger Gilchrist, 
Ulli and many other fine programmers that I don’t have a record of.
Thank you to all these generous people! 

A Special Thank You to Planet Source Code for
all the support given me all this years.
Thank Very Much PSC! 

Two DLLs from Info-Zip will be needed and you can get them from
the following links;
ftp://ftp.info-zip.org/pub/infozip/WIN32/zip231dN.Zip
Extract the DLL [ZIP32.DLL] and rename it as [Info231-zip32vc.dll]
ftp://ftp.info-zip.org/pub/infozip/WIN32/unz552dN.Zip
Extract the DLL [UNZIP32.DLL] and rename it as [Info552-unzip32vc.dll]
In case there is a problem getting the files from Info-Zip download them from
http://www.vccs.info/misc/infozipdlls.zip
Both DLLs are in a ZIP file and renamed.
Copy the DLLs to the System32 directory.
If your system fails to find the DLLs then move them to the application directory.


Quick Start.

Create two directories on your hard disk with the following names;
[PSC CodeOf TheDay] and [PSC SavedCode].
These directories can be any name. This is just an example.
Drag all, or some, [Code of the Day] messages and drop in the
[PSC CodeOfTheDay] directory.
This operation can be a problem on Win9x. You’ll have to work it out.
The emails must be Plain Text.
If you don’t have Code of the Day messages then skip this.
It is needed just to get the Category Names.
 
Go to Planet Source Code and select some of your favourite code pages
and save the ZIP file and page to the [PSC SavedCode] directory.

Run [PSC FileStore] program, go to 
[Tools/Code of the Day Categories/Import Code of the Day]
Click [Import Emails] button and navigate to [PSC CodeOfTheDay] directory
All the messages will be scanned to import some details.
If you don’t have [Code of the Day] messages skip this.

Go to [File/ Import Zip Files] and click on [Select Folder] button.
Point to the [PSC SavedCode] directory.
Will showed to you how many ZIP files found.
The program can backup your files before the import starts.
The default is [YES] but it can be changed to [NO].  
Click the [Start Import] button.
The import will start and the information will be displayed to you.
[No of Files Processed]; How many files were successfully imported
[No of Files Skipped]; How many files were skipped. 
The file [@PSC_ReadMe_???_?.txt] or [File_Id.Diz] was not found.
I have been using [File_Id.Diz] before PSC came with a grate idea of   [PSC_ReadMe_???_?.txt]
[No of Duplicate Files]; How many files are duplicated.
These files will be moved to a subdirectory of [PSC SavedCode]
[No of Files with Errors]; How many files couldn’t be Unziped.
These files will be moved to a subdirectory of [PSC SavedCode]
There is a problem here.
If Info-Zip detects the ZIP file as corrupted or invalid.
Manually Unzip the file and Zip it.
Now import the file again, most of the times works.  

Go to [Maintenance/ Edit and Moving Files]
Click the [Categorize] button.
All Files/Titles will be moved from the default import directory
to the proper Category except if the Title doesn’t exist in the [Code of the Day].

The remaining files, if no [Code of the Day] will be all files, in the
default category can be moved to a Category of your choice.
Go to [Maintenance/ Categories - Add, Edit and Delete]
to create new Categories and return here.
Click the button [Load] select the import default category.
I name it as [### Imported Files To Be Moved].
After all Titles have been loaded to the list select one Title
click the [Move] button and select the Category to be moved to.

That’s it!
Just a very short explanation.

For more information read the Tool Tips.

If you have any questions please send me a email at PSC
or email me at   jginacio@hotmail.com

I hope you like and enjoy the code!

George 
