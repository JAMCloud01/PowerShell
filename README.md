 ---------------------------------------
| Microsoft 365 User Offboarding and Management Automation using PowerShell |
 ---------------------------------------

Small script that operates under a while loop that takes advantage of using two import modules: Microsoft Graph and ExchangeOnline. You are prompted to make a selection between 1-5 with each having their own functionality.

1) Iterates through each employee looking for active inbox rules tied to the employee. If nothing is found, the script will make a mention of it.

2) Utilizes Out-GridView to view a list of employees, depending on the amount selected, the script iterates through each employee looking for active licenses found within disabled accounts. If licenses are found, you are prompted to give Y/N to remove.

3) Utilizes Out-GridView similiar to the above. Option 3 prompts the user on whom they want to select to convert their mailbox to shared.

4) Much similiar to option 3 but instead converts the user selected back to regular operating mailbox.

5) Here the option is given to revoke the selected user's session. Depending on how many are selected.
