# skytem-linebiblebuddy

WHAT THIS PROGRAM DOES
Basically it will fill the block sheet in the linebible from the linereports.
- Using the program first choose the folder with your linereports in. 
- You need ALL line reports for a given block sheet in that folder.
- the program will make a linebible_copy when it is run. If the program is being run 2 times in a row, the backup file will be overwritten by the new backup file.
So fx. if you only want to fill block3 fields in the linebible, just keep all linereports for that block in a folder fx. called block3.
The program will attempt to fill all sheets with matching numbers and will automatically exclude the 920000-930000.


YOU NEED PYTHON IN ORDER TO RUN LINEBIBLEBUDDY - TO GET IT DO FOLLOWING STEPS

1. use microsoft store to install the newest python version avaliable.

2. launch a terminal by clicking windows key and write "cmd" and hit enter

3. verify python version with "python -V"

4. navigate to the folder containing requirements.txt. ("cd c:\path\to\txt")

5. type "pip install -r requirements.txt". This will install neccesary modules for linebiblebuddy to run.

5.1 your will most likely get a path error - add the paths by
5.2 OPTION1
pressing windows key -> "edit environmental variables for your account" -> "new" and add the paths specified 
when running pip install.
5.2 OPTION2
Use CLI and fill this : setx NEW_VAR "C:\NewPath" (I haven't tried this but chatGPT says its it is do :)

6. run "pip install -r requirements.txt" in terminal window again - this time you should have no errors.

if python has been installed correctly on your windows machine, you should now be able to double click the linebiblebuddy.py
to execute it.
