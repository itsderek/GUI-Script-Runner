# GUI-Script-Runner

## Summary:

Instead of manually running the plethora of SQL scripts each month I created an automated way to run them using Python. This version of the program now includes a GUI which allows you to select more than one billing system to run at a time. The GUI was created using PySimpleGUI and was made with the intent to distribute among other members of my organization who aren't as comfortable with the command prompt.

The GUI utilizes threads to run the SQL scripts which means the interface is non-blocking. During testing I used SQLAlchemy along with an excel file to simulate a database, the following article was very helpful:
[Creating a Mock Database](https://blog.devgenius.io/creating-a-mock-database-for-unittesting-in-python-is-easier-than-you-think-c458e747224b)

Overall this program should be a useful tool for my team members and expedite the script running process for month end close.

## Billing System Selection Screen:

![First Screen](https://github.com/itsderek/GUI-Script-Runner/blob/main/screenshots/first_screen.PNG?raw=TRUE)

## Script Running Screen:

![Second Screen](https://github.com/itsderek/GUI-Script-Runner/blob/main/screenshots/second_screen.PNG?raw=TRUE)
