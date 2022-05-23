# ReminderApp
Google Sheets reminder application that automatically sends email reminders after specified number of days

The objective of this project is to create a list of reminders on Google Sheets that will be received via email once a specified number of days 
have passed. It can be used as a personal diary or as an organization-wide scheduling tool. Users can activate admin priviledges where only
specific users can activate/deactivate email alerts. Since the app requires admin email access, it will be flagged by Gmail as unsafe. Triggers
can be customized to run hourly, daily, weekly, bi-weekly etc.

Rather than having a winding list of reminders, the app also creates a new workbook for each month and notifies all users (admins & editors). All
active triggers are by default deleted 30 days after the last reminder in the workbook.


![sample](https://user-images.githubusercontent.com/46036415/169867022-1ab0947b-ef9c-42ed-b978-354d450db1b8.png)
