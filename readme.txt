A python script that updates the temperatures of the given cities in the excel sheet every 10 seconds. 
Python libraries used- openpyxl, requests, json, time
Weather API- openweathermap.org

Additional info for running the application correctly-
  openpyxl library is used, install it using pip using the command in cmd "pip install openpyxl"
  run the index.py file with the weather.xlsx file in the same folder.
  keep the excel file closed while running the script.
  the script will update the temperatures every 10 seconds
  use ctrl+c to exit out of the script any time.

More info on customizing the data-
  Changing C/F will make it switch between updating celsius and fahrenhiet
  Changing 1/0 will toggle the temperature being updated or not.
  Make these changes when the script IS NOT running.
  Save the changes in excel file and exit out. Then run the script again.
