# Excel-Smart-Search-App
Application for comparing strings between two excel files. Developed for CERN TE-CRG-IC.

<!-- PROJECT LOGO -->
<br />
<p align="center">
  <a href="https://github.com/othneildrew/Best-README-Template">
    <img src="app.png" alt="Logo" width="555" height="380">
  </a>
  <h3 align="center">Excel Search App GUI</h3>
</p>


<!-- ABOUT THE PROJECT -->
## About The Project
The repository includes the installation (.exe) file and the source code of the application.
The application was implemented for TE-CRG-IC in order to search for existing components in stores, when a new design is about to start.
  
This implementation uses:
- customtkinter library for GUI
- openpyxl library for working with excel files    
- smart fuzzywuzzy string search algorithm


<!-- GETTING STARTED -->
## Getting Started
To get started you simply have to run the installation file and complete the installation. 

### User manual
The reason for the creation of this app was to automate the procedure of searching components in stores (on EAM). 
- [cmake](https://cmake.org/download/)
- [boost](https://www.boost.org/users/download/) library, which can also be extracted from the 7z included
  - Make sure to include the path of the library in the CMakeLists
- A compiler that supports std11 and threading
  - (This project was tested with [mingw64](http://mingw-w64.org/doku.php) version 8)

### Modifing the source code
PyCharm was used for the creation of this app. In order to run the python script you need to have the following libs installed: 
<br />
<p align="center">
  <a href="https://github.com/othneildrew/Best-README-Template">
    <img src="libs.png" alt="Logo" width="555" height="380">
  </a>
  <h3 align="center">Libraries</h3>
</p>
You also have to manually replace the original library folder 'customtkinter' using the folder from this repo to your site-packages file because it has some minor modifications. Make sure all necessary gifs and photos are included. 

### Converting py to exe
- Make sure that all needed libraries are installed to your computer
- Install auto_py_to_exe application
- Open cmd and run: python -m auto_py_to_exe
- Define the main.py script location
- Select --> One Directory and Window Based (hide the console) options
- Select the 'custontkinter' location folder as an additional file
- All libraries, logos and gifs should be selected as additional files. You can simply include a folder which contains all of them
- Define output directory from settings
- Convert!

### Converting exe to installation file
You can easily do this by using [NSIS](https://nsis.sourceforge.io/Download) application. Simply convert the produced from auto_py_to_exe folder to zip and use the NSIS app. If you wish to get better - more beautiful result you can use [Advanced installer](https://www.advancedinstaller.com/). It also has an option to produce a shortcut to the desktop after installation. 

