# Improved Excel Converter

## Rationale 
We use an excel converter to take Excel files from clients and put them into a form that works with internal scripts. The current Excel converter in use is very good, but several years out of date. Any conversions requiring something more complex than changing the file extension aren't really covered by the converter. 

The conversion to .csv additionally will not support multi-page client files, as Excel only allows one worksheet per CSV. This means only the last sheet in a multi-page .xlsx or .xls file would survive, potentially excising important customer data. 

## The Structure 

The python script combines two powerful libraries to create an efficient file conversion process with the potential to be extended in many ways: watchdog (a system event monitor) and pandas (a data frame handling library). 

The system monitor watches for the creation of files in the "input" directory, then processes them dependent on their type. 

## The Build

The build was done with Pyinstaller and removes the need to implement a Python interpreter on the endpoint the script is used on. Updates to source code will not be built into the executable until compiled again after pull, so ensure that this is done before updating in production. I will attempt to build each new major release and tag, so this probably won't be necessary in each case. 

The build is located in the dist/ directory. It will create everything necessary on first run or you can create an "input", "output", and "manrev" folder. It will handle files already being in the input folder as well, so this doesn't need to be done while the program is running. 

## Dependencies

Dependencies are only really necessary for running from source; they can be found in requirements.txt. Using the build requires only the permission to run executables. The build is made for Windows x64 systems and hasn't been tested elsewhere. 


## Contributing

If you notice a missing feature, a bug, or an improvement, please submit a pull request (PR).  

**Guidelines:**
- Fork the repository and create a new branch for your changes.
- Follow the existing code style.
- Write tests for new features or bug fixes if applicable. 
- Submit a clear PR with a description of what your changes do. 

I will try to review as soon as possible. Poke me if I forget. 

## Extensibility

The program's flow of taking items placed in a folder, whether when it's running or not, could easily be extended to add different processing (e.g., encryption, alteration, compression), with minimal changes to the source code. I hope that this can be used for a variety of purposes and give users an easy jump ahead of wrangling with the initial balancing act between os and watchdog. 

The other watchdog functions are included in the source, commented out. This will allow anybody to uncomment and experiment with the length the program can go to.

## License

Copyright (C) 2025 Ryan Piazza

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.