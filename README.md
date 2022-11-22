# Excel-Smart-Search-App
Application for comparing strings between two excel files. Developed for CERN TE-CRG-IC.

<!-- PROJECT LOGO -->
<br />
<p align="center">
  <a href="https://github.com/othneildrew/Best-README-Template">
    <img src="app.png" alt="Logo" width="555" height="380">
  </a>
  <h3 align="center">Excel Search App</h3>
</p>


<!-- ABOUT THE PROJECT -->
## About The Project
The repository includes the installation (.exe) file and the source code of the application.
The application was implemented for TE-CRG-IC in order to for search existing components in stores, when a new design is about to start.
  
This implementation uses:
- customtkinter library for GUI
- openpyxl library for working with excel files    
- smart fuzzywuzzy string search algorithm


<!-- GETTING STARTED -->
## Getting Started
To get started you simply have to run the installation file and complete the installation. The app will be installed and fully functional. 

### Prerequisites
To compile this implementation of RSA you will need:
- [cmake](https://cmake.org/download/)
- [boost](https://www.boost.org/users/download/) library, which can also be extracted from the 7z included
  - Make sure to include the path of the library in the CMakeLists
- A compiler that supports std11 and threading
  - (This project was tested with [mingw64](http://mingw-w64.org/doku.php) version 8)

### Installation
To run the example: 
- Open the example folder and type:
```
mkdir build
cd build
cmake ..
make
```
- Add a file named plain.txt with the message you want to encrypt before running the output file.

To run the tests, follow the same procedure.

