#CV Creator

This is a Python application that allows you to create, save, load, and export your CV. It uses the Tkinter library for the GUI and the python-docx library to export the CV to a Word document.
Features

    Personal Information: Enter your name, email, cellphone number, LinkedIn profile, ID number, driver's license code, citizenship, and passport number.
    Experience: Add multiple experiences with job title, company, period, and description.
    Education: Add multiple education entries with school, degree, major, and dates.
    Skills: Add multiple skills with type and description.
    Awards: Add multiple awards with award name, awarded by, and date.
    Certificates: Add multiple certificates with title, issued by, and date.
    Photo: Upload a photo to be included in your CV.
    Save and Load: Save your CV data to a JSON file and load it back into the application.
    Export: Export your CV to a Word document.

Dependencies

    Python 3
    Tkinter
    python-docx
    PIL

How to Run

    Install the dependencies. If you're using pip, you can do this by running pip install python-docx pillow.
    Run the script with Python 3 by typing python cv_creator.py in your terminal.

Code Structure

The code is structured into several functions, each handling a specific part of the application:

    main(): The main function that sets up the GUI and starts the Tkinter event loop.
    add_experience(), add_education(), add_skill(), add_award(), add_certificate(): Functions to add entries to the respective sections.
    remove_entry(): Function to remove an entry from a section.
    save_data(): Function to save the CV data to a JSON file.
    load_data(): Function to load the CV data from a JSON file.
    export_to_docx(): Function to export the CV data to a Word document.
    collect_cv_data(): Function to collect the CV data from the GUI.

Note

This application does not validate all the input data. Please ensure that the data you enter is correct and formatted as you want it to appear in the CV.
