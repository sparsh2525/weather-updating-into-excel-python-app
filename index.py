import requests, json , openpyxl, time

cities=['London','New York','Paris','Tokyo','Moscow','Dubai','Singapore','Barcelona','Los Angeles', 'Delhi']  

api_key = "86fd4d9ab1324ecf21b78e2f96122c52"
base_url = "http://api.openweathermap.org/data/2.5/weather?"
print('OPENPYXL LIBRARY IS USED')
print('USE CTRL+C TO EXIT OUT OF THE SCRIPT')
print('THIS SCRIPT UPDATES THE TEMPERATURE OF THE CITIES EVERY 10 SECONDS')
print('KEEP THE EXCEL FILE CLOSED WHILE RUNNING THE SCRIPT\n')

while 1:
    try:
        wb = openpyxl.load_workbook('weather.xlsx') 
        sheet = wb["Sheet1"]
        cells=sheet['A1':'D10']
        for city,temp,cf,yn in cells:
            if yn.value==1:
                complete_url = base_url + "appid=" + api_key + "&q=" + city.value
                response = requests.get(complete_url) 
                x = response.json()
                if x["cod"] != "404":  
                    t = x["main"]["temp"] 
                    if cf.value=='C':
                        print(f'Updating the temperature in celsius for {city.value}')
                        temp.value=t-273.15
                    elif cf.value=='F':
                        print(f'Updating the temperature in fahrenhiet for {city.value}')
                        temp.value=((t-273.15) * 9/5) + 32
            else:
                print(f'Not updating the temperature for {city.value}')
                temp.value='Not updating'
                
        wb.save("weather.xlsx")
        print('Next update in 10 seconds.\n')
        time.sleep(10)

    except:
        exit()