# MAS-Weekly-Average-Exchange-Rate-API

Step 1: MAS API Subscription
Subscribed to the MAS API using my company email. After successfully subscribing to the MAS weekly exchange rate API, I received the API key required to retrieve the data. 
MAS API Portal: https://eservices.mas.gov.sg/apimg-portal/home
 

Step 2: Python Script Development
Using this API, I created a Python script to pull the weekly exchange rate data and export it into an Excel file containing only our required currencies.
Below is my python script snippet, (hide API key for credential purpose) 
 
Step 3: Automation Using Task Scheduler
To automate the process, I used Windows Task Scheduler on my laptop and set the script to run every Friday at 3:00 PM.


Error Prevention Logic 
During testing, I noted that MAS normally updates the weekly average exchange rate every Friday at around 3:00 PM. 
However, if Friday is a public holiday, the data may be updated earlier, such as on Thursday. 
Since the API only returns data for available weekly dates, calling the API with a date that has no published data will return an empty result.
To handle this, I added logic in the Python script to use today’s date and automatically check back up to the previous 7 days. 
Once the script finds the latest available MAS exchange rate data, it will retrieve the result and export it into Excel. 
This helps prevent errors when the scheduled run date does not exactly match the MAS published date.

Current Limitation
One current limitation is that the Task Scheduler is set up on my laptop, so my laptop must be on at the scheduled time for the automation to run.

