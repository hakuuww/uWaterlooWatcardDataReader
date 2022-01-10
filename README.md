# uWaterlooWatcardDataReader
The purpose of this project is to read transaction and balance data from a UWaterloo Watcard account and inputting the read data into an excel file
The project is coded in python and used Selenium&Openpyxl framework
columnar and click libraries were also used to format data in run window


the following must be completed before running the main.py file:
  1. Have Python installed and setted up

  2. In command line install the following by typing(one by one):
     pip install selenium
     pip install columnar
     pip install click
     pip install openpyxl
     
  3. Have Chrome or Firefox installed and download corresponding driver.exe files
    The proccess for Chrome is demonstated below:
    In a seperate tab, type in chrome://version into the search bar
    
    chrome://version
    
   The current version of the chrome browser will be shown right beside "Google Chrome:"
   Go to https://chromedriver.chromium.org/downloads
   
    https://chromedriver.chromium.org/downloads
       
   Locate the corresponding folder that matches your Chrome version
       Download the zip file correspoinding to your operating system
       Create an empty folder in the C drive called SeleniumDrivers
       Extract from the zip file and move the exe file into the created folder (you can move it into any folder but you also have to change the executable_path variable at           the begining of the python script)
      
   The chromedriver.exe file should now be at the location " C:/SeleniumDrivers/chromedriver.exe "
   
    C:/SeleniumDrivers/chromedriver.exe
       
   You are now all set for running the main.py program
   
After running the program, an Excel file will be generated inside the same folder with the name: "WatcardTransactions.xlsx"
   You can open the Excel file and format the data by 
      selecting all cells with Ctrl+A
      pressing Alt+H then O then I
       
        Ctrl+A
        Alt+H then O then I
      
   To display all data by AutoFitting column width for all cells
   
   Enjoy your Watcard data!
       
       
       
       
       
       
       
       
       
      
     
     
     
 
    

