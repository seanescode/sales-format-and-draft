# Sales Report Automation

### What the automation does
I created an automation to make the process of preparing and emailing a weekly sales report quicker and more consistent.

<img width="1212" height="438" alt="report-auto-formatter" src="https://github.com/user-attachments/assets/e6763fd2-e152-4c7d-9bd7-d7568727b208" />

The automation:

1. Opens the Excel sales report in the background.
2. Creates summary tables from the sales data.
3. Formats the report and saves the completed version.
4. Creates an Outlook email, fills in the email details and attaches the completed report.

The user can then review the email before sending it.

### The client can change the settings

The automation has a separate configuration file where the client can change certain settings without changing the Python code.

For example, they can change:

- **Report formatting** — such as colours, headings and bold formatting.
- **Email details** — such as who the email is sent to, who is CC'd, the subject and the email body.

This means the automation can be adjusted as the client's requirements change without needing to modify the main program.

### The report can handle changes

The summary tables are not placed in fixed cells.

Instead, the automation looks at where the report data ends and places the summary tables to the right of it.

For example, if the original report gains additional columns, the summary tables can move further to the right rather than being placed over the new data.

This makes the automation more flexible when the layout of the report changes.

### What is needed to run it

- Windows
- Microsoft Excel desktop
- Classic Outlook desktop

The current version uses Pywin32 and Xlwings to interact with these applications.

### Why I chose this approach

I chose a desktop-based approach for this version because it allows the automation to work directly with the Excel and Outlook applications that a client may already be using.

It also avoids requiring this version of the automation to connect to Microsoft Graph.

There are trade-offs, however. Because this version relies on desktop applications, it is more dependent on the client's computer setup and is not the approach I would necessarily choose for a larger, cloud-based automation.

### How I would develop it further

If this automation needed to be used by more people or deployed on a larger scale, I would look at reducing its dependency on locally installed Microsoft Office applications.

For example, I could:

- Use Microsoft Graph to interact with Outlook without relying on the Outlook desktop application.
- Reduce the number of Microsoft-specific dependencies.
- Separate the report-processing part of the automation from the Excel and Outlook parts.
- Move towards a solution that can run centrally rather than on an individual user's computer.

The current version is therefore a practical solution for a Windows-based Microsoft Office environment, while also providing a foundation that could be developed into a more scalable solution if the requirements increased.
