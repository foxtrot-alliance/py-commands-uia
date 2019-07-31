# py-commands-uia

This program allows you to work with advanced GUI automation based on either win32 or uia backend technology. You can run the program via the CMD or as part of an automation script in an RPA tool like Foxtrot. This solution is meant to supplement Foxtrots core functionality and enable you to interact very precisely with elements and objects that the standard Foxtrot targeting technology might not be able to detect. For example, if you have an application where Foxtrot is limited in its ability to target the individual buttons and fields, this solution is very useful for you. The solution is written in Python using the module "pywinauto". You can see the [full source code here](https://github.com/foxtrot-alliance/py-commands-uia/blob/master/py-commands-uia.py).

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

## Commands

All the available commands are specified [here](#all-available-parameters). Note, all parameters surrounded by [-x "X"] means that they are optional. For a more detailed description of each command, read the [detailed command description section](#detailed-command-description).

For each command, there are three things you need to consider:
1. What technology/backend do you intend to use?
    1. "win32"  : Appropriate for old applications such as MFC, VB6, VCL, simple WinForms controls and most of the old legacy apps
    2. "uia"    : [Default] Appropriate for most applications, especially WinForms, WPF, Store apps, Qt5, browsers
```
-backend "X"
```
2. What is your target to engage with? Here, you will need to specify a logic to locate the element similar to automating in web browsers with HTML.
```
-app "X" -main_window "X" [-child_window1 "X"] [-child_window2 "X"] [-child_window3 "X"] [-child_window4 "X"] [-child_window5 "X"]
```
3. What action do you wish to perform?
    1. "PRINT"
    2. "PRINT_FILE"
    3. "CLICK"
    4. "DOUBLECLICK"
    5. "RIGHTCLICK"
    6. "SEND"
    7. "SELECT"
    8. "LOCATION"
    9. "WAIT"
```
-command "X" [-value "X"]
```
A full command could look like this:
```
PROGRAM_EXE_PATH -backend "X" -app "X" -main_window "X" -child_window1 "X" -command "X"
```
Example:
```
PROGRAM_EXE_PATH -backend "uia" -app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Maximize, control_type=Button" -command "click"
```
The solution will give an output to the selected variable in the DOS Command action to indicate whether the command was executed successfully or not.

## Finding the target

To find the target, there is a [long list of parameters](https://pywinauto.readthedocs.io/en/latest/code/pywinauto.findwindows.html#pywinauto.findwindows.find_elements) you can use:
* name: The name of the element
* class_name: Elements with this window class
* class_name_re: Elements whose class matches this regular expression
* path: The path used to launch the application
* process: Elements running in this process
* title: Elements with this text
* title_re: Elements whose text matches this regular expression
* control_id: Elements with this control id
* control_type: Elements with this control type (string; for UIAutomation elements)
* auto_id: Elements with this automation id (for UIAutomation elements)
* framework_id: Elements with this framework id (for UIAutomation elements)

Here are some examples of how you could find the whole window of a newly opened Notepad window:
```
-app "title=Untitled - Notepad" -main_window "title=Untitled - Notepad"
-app "path=notepad.exe" -main_window "title_re=.* Notepad"
```
Here is an example of how to find the maximize button of Notepad:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Maximize, control_type=Button"
```

### Printing all elements

The tough part is figuring out how to find the parameters to find the element. For this, you can use the "PRINT" command. The "print" command is going to locate the specified target and print all the controls within the specified target. This is what you will use in order to find the actual target that you need to engage with.

For Notepad, you can, for example, run this command:
```
PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -command "print"
```
The output of this command will be something like:
```
Control Identifiers:

Dialog - 'Untitled - Notepad'    (L1349, T337, R2459, B821)
['Untitled - NotepadDialog', 'Untitled - Notepad', 'Dialog']
child_window(title="Untitled - Notepad", control_type="Window")
   | 
   | Edit - 'Text Editor'    (L1357, T388, R2451, B813)
   | ['', 'Edit', '0', '1', 'Edit0', 'Edit1']
   | child_window(title="Text Editor", auto_id="15", control_type="Edit")
   |    | 
   |    | ScrollBar - 'Vertical'    (L2434, T388, R2451, B813)
   |    | ['ScrollBar', 'Vertical', 'VerticalScrollBar']
   |    | child_window(title="Vertical", auto_id="NonClientVerticalScrollBar", control_type="ScrollBar")
   |    |    | 
   |    |    | Button - 'Line up'    (L2434, T388, R2451, B405)
   |    |    | ['Button', 'Line upButton', 'Line up', 'Button0', 'Button1']
   |    |    | child_window(title="Line up", auto_id="UpButton", control_type="Button")
   |    |    | 
   |    |    | Button - 'Line down'    (L2434, T796, R2451, B813)
   |    |    | ['Button2', 'Line downButton', 'Line down']
   |    |    | child_window(title="Line down", auto_id="DownButton", control_type="Button")
   | 
   | TitleBar - ''    (L1373, T340, R2451, B368)
   | ['2', 'TitleBar']
   |    | 
   |    | Menu - 'System'    (L1357, T345, R1379, B367)
   |    | ['SystemMenu', 'System', 'Menu', 'System0', 'System1', 'Menu0', 'Menu1']
   |    | child_window(title="System", auto_id="MenuBar", control_type="MenuBar")
   |    |    | 
   |    |    | MenuItem - 'System'    (L1357, T345, R1379, B367)
   |    |    | ['System2', 'SystemMenuItem', 'MenuItem', 'MenuItem0', 'MenuItem1']
   |    |    | child_window(title="System", control_type="MenuItem")
   |    | 
   |    | Button - 'Minimize'    (L2312, T338, R2359, B368)
   |    | ['Button3', 'MinimizeButton', 'Minimize']
   |    | child_window(title="Minimize", control_type="Button")
   |    | 
   |    | Button - 'Maximize'    (L2359, T338, R2405, B368)
   |    | ['Button4', 'MaximizeButton', 'Maximize']
   |    | child_window(title="Maximize", control_type="Button")
   |    | 
   |    | Button - 'Close'    (L2405, T338, R2452, B368)
   |    | ['Button5', 'CloseButton', 'Close']
   |    | child_window(title="Close", control_type="Button")
   | 
   | Menu - 'Application'    (L1357, T368, R2451, B387)
   | ['ApplicationMenu', 'Application', 'Menu2']
   | child_window(title="Application", auto_id="MenuBar", control_type="MenuBar")
   |    | 
   |    | MenuItem - 'File'    (L1357, T368, R1389, B387)
   |    | ['FileMenuItem', 'File', 'MenuItem2']
   |    | child_window(title="File", control_type="MenuItem")
   |    | 
   |    | MenuItem - 'Edit'    (L1389, T368, R1423, B387)
   |    | ['EditMenuItem', 'Edit2', 'MenuItem3']
   |    | child_window(title="Edit", control_type="MenuItem")
   |    | 
   |    | MenuItem - 'Format'    (L1423, T368, R1475, B387)
   |    | ['Format', 'FormatMenuItem', 'MenuItem4']
   |    | child_window(title="Format", control_type="MenuItem")
   |    | 
   |    | MenuItem - 'View'    (L1475, T368, R1514, B387)
   |    | ['ViewMenuItem', 'MenuItem5', 'View']
   |    | child_window(title="View", control_type="MenuItem")
   |    | 
   |    | MenuItem - 'Help'    (L1514, T368, R1553, B387)
   |    | ['HelpMenuItem', 'Help', 'MenuItem6']
   |    | child_window(title="Help", control_type="MenuItem")
SUCCESS
```
This finds the Notepad application and the main window to print all the elements inside the window. Now, using this information, you can engage with, for example, the maximize button like mentioned earlier:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Maximize, control_type=Button"
```
TIP: Notice that the use of quotation marks is slightly different in the command line compared to the print output.
You could also find the maximize button using the name of the element like this:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "name=MaximizeButton"
```
If you need to engage with child elements, elements inside other elements, which is quite typical, such as the application menu options, you can engage with the "File" button in the menu either directly or via the menu:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=File, control_type=MenuItem"
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Application, auto_id=MenuBar, control_type=MenuBar" -child_window2 "title=File, control_type=MenuItem"
```
### Printing to files

For applications with a lot of elements, it can be more convenient to print the output to a file instead of the console. To do this, you have to include "file" in the command and specify the folder using the "value" parameter like this:
```
PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -command "print_file" -value "c:\"
```
This will generate a text file called "print_control_identifiers.txt" in the specified folder.

## Multiple commands

In some cases, you need to execute multiple commands in one sequence, for example, if you need to click in a menu to expand a dropdown of options to then click on one of the options. To show an example, here we click to expand the Notepad menu to then click to open the Page Setup in Notepad + clicks to select the "Landscape" orientation.
```
PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=File, control_type=MenuItem" -command "click" -hover "true" && PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=File, control_type=MenuItem" -child_window2 "title=Page Setup..., control_type=MenuItem" -command "click" && PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Page Setup, control_type=Window" -child_window2 "title=Landscape, control_type=RadioButton" -command "click"
```
This single command is actually three commands put together in a sequence using "&&". These are the three commands separated:
```
PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=File, control_type=MenuItem" -command "click" -hover "true"
PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=File, control_type=MenuItem" -child_window2 "title=Page Setup..., control_type=MenuItem" -command "click"
PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Page Setup, control_type=Window" -child_window2 "title=Landscape, control_type=RadioButton" -command "click"
```
This is also relevant if you need to find out how to access elements only visible while hovering over an item. In case you need to figure out how to engage with the elements in the Notepad menu, you need to click on the menu item and immediately after print information. You can use the same approach for this:
```
PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=File, control_type=MenuItem" -command "click" -hover "true" && PROGRAM_EXE_PATH -app "path=notepad.exe" -main_window "title_re=.* Notepad" -command "print"
```
Notice that in the second command, the print command, we look in the whole window, not only in the menu. This way you make sure to get information about all elements in the application after the click.

## All available parameters
```
-backend: "win32"/"uia", default = "uia"
  This will determine the technology that the program will utilize.

-app: the parameters to connect to the application, required
  This will determine how the program will find the application to engage with.

-main_window: the parameters to connect to the window of the application, required
  This will determine how the program will find window of the application.

-child_windowX: the parameters to locate the target element, you can define between 0-5 child windows
  This is the (optional) additional layers to locate the element to engage with.

-command: "print"/"print_file"/"click"/"doubleclick"/"rightclick"/"send"/"select"/"location"/"wait", required
  This is the command to perform with the located element.

-value: "X", required if "print_file"/"send"/"select" commands are used
  This will instruct how to perform the specified command if relevant.

-hover: "true/false", default = "false"
  If a click command is selected, this determines whether the mouse should remain in the position after clicking.

-wait: "X", required if the "wait" command is used
  This will instruct the program on how many seconds it should wait.

-traces: "true"/"false", default = "false"
  This determines whether you wish the output to include traces, information about the execution.
```

## Detailed command description

### Print
Parameters:
```
-app "X" -main_window "X" [-child_window1-5 "X"] -command "print"
```
Examples:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -command "print"
```

### Print to file
Parameters:
```
-app "X" -main_window "X" [-child_window1-5 "X"] -command "print_file" -value "X"
```
Examples:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -command "print_file" -value "c:\"
```

### Click
Parameters:
```
-app "X" -main_window "X" [-child_window1-5 "X"] -command "click" [-hover "X"]
```
Examples:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Maximize, control_type=Button" -command "click"
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=File, control_type=MenuItem" -command "click" -hover "true"
```

### Double-click
Parameters:
```
-app "X" -main_window "X" [-child_window1-5 "X"] -command "doubleclick" [-hover "X"]
```
Examples:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Maximize, control_type=Button" -command "doubleclick"
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=File, control_type=MenuItem" -command "doubleclick" -hover "true"
```

### Right-click
Parameters:
```
-app "X" -main_window "X" [-child_window1-5 "X"] -command "rightclick" [-hover "X"]
```
Examples:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Maximize, control_type=Button" -command "rightclick" -hover "true"
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Text Editor, control_type=Editor" -command "rightclick" -hover "true"
```

### Send
NOTE: You can send keystrokes using [these key codes](https://pywinauto.readthedocs.io/en/latest/code/pywinauto.keyboard.html).

Parameters:
```
-app "X" -main_window "X" [-child_window1-5 "X"] -command "send" -value "X"
```
Examples:
```
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Text Editor, control_type=Editor" -command "send" -value "Hello World!"
-app "path=notepad.exe" -main_window "title_re=.* Notepad" -child_window1 "title=Text Editor, control_type=Editor" -command "send" -value "Hello{ENTER 2}World!"
```

### Select
Parameters:
```
-app "X" -main_window "X" [-child_window1-5 "X"] -command "select" -value "X"
```

### Location
Parameters:
```
-app "X" -main_window "X" [-child_window1-5 "X"] -command "location"
```

### Wait
Parameters:
```
-app "X" -main_window "X" [-child_window1-5 "X"] -command "wait" -wait "X"
```
