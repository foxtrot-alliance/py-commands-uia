# py-commands-uia

This program allows you to work with advanced GUI automation based on either win32 or uia backend technology. You can run the program via the CMD or as part of an automation script in an RPA tool like Foxtrot. This solution is meant to supplement Foxtrots core functionality and enable you interact very precisely with elements and objects that the standard Foxtrot targeting technology might not be able to detect. For example, if you have an application where Foxtrot is limited in its ability to target the individual buttons and fields, this solution is very useful for you. The solution is written in Python using the module "pywinauto". You can see the [full source code here](https://github.com/foxtrot-alliance/py-commands-uia/blob/master/py-commands-uia.py).

## Installation

1. Download the [latest version](https://github.com/foxtrot-alliance/py-commands-uia/releases/download/v0.0.1/py-commands-uia_v0.0.1.zip).
2. Unzip the folder somewhere appropriate, we suggest directly on the C: drive for easier access. So, your path would be similar to "C:\py-commands-uia_v0.0.1".
3. After unzipping the files, you are now ready to use the program. The only file you will have to be concerned about is the actual .exe file in the folder, however, all the other files are required for the solution to run properly.
4. Open Foxtrot (or any other RPA tool) to set up your action. In Foxtrot, you can utilize the functionality of the program via the DOS Command action (alternatively, the Powershell action).

## Usage

When using the program via Foxtrot, the CMD, or any other RPA tool, you need to reference the path to the program exe file. If you placed the program directly on your C: drive as recommended, the path to your program will be similar to: 
```
C:\py-commands-uia_v0.0.1\py-commands-uia_v0.0.1.exe
```
TIP: Make sure NOT to surround the path with quotation marks in your commands.
