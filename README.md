# Month-End-Close-De-Beers

Project Title:
Month-End-Close-De-Beers

Description:
The purpose of the process is to close month end on SAP for De Beers SA. Currently the process is carried out manually where the user copies data from an Excel "closure" template to SAP. For De Beers SA alone this is over 550 lines of data which cannot be copied over to SAP at once. The data can also not be uploaded to SAP. This, coupled with the fact that this process usually occurs at night (when all transactions have been processed) opens room for errors to occur. A further impact is that while the team is working on capturing this data on SAP, other people cannot access the company codes causing further delays in other areas. The process of closing the month end for De Beers SA is very manual and therefore time-consuming and prone to errors.

This process was automated, which enhanced accuracy, controls, and saved time.

# Prerequisites and Dependencies:
To proceed with the task, we will need the following files:
1. Annual Calendar
2. Valid Company Codes

# Code Explanation:
The script is for automating some tasks related to SAP GUI interaction, specifically for the "Month-End Close De Beers" application. Here's a breakdown of what the code does:
1. It starts by checking if the script version matches the expected version maintained by the Automation Team. If it doesn't match, it prompts the user to contact the team and then exits the script.
2. It prompts the user to select a file containing company codes.
3. It reads the Excel file containing the company codes and stores them in a DataFrame.
4. It opens SAP GUI using subprocess.Popen and connects to the SAP GUI scripting engine using the win32com library.
5. It prompts the user to enter the SAP server name and based on the server name, it performs different actions. For instance, if the server name is "QP8", it opens a connection and creates a session to execute transactions related to the "Month-End Close De Beers" application.
6. It defines two functions run_first_loop and run_second_loop, which seem to handle different scenarios based on SAP responses. These functions seem to fill a DataFrame with data retrieved from SAP GUI tables.
7. Within these functions, there's a loop that interacts with SAP GUI elements to extract data and populate the DataFrame.
8. Finally, it writes the DataFrame to an Excel file.
Overall, the script automates the process of fetching data from SAP GUI tables based on certain criteria and writes the results to an Excel file.
