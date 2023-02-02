# Python_WebScrapping
These is one of the web scrapping projects I have done using Python.
## OU Results Scrapping
**This Project has an inbuilt GUI (Tkinter) which is used to take multiple input values -**   
Chrome Web Driver Location  
Input File Location  
Output File Location  
First URL for the results page  
Secoung URL is there is a re-evaluation page  
  
**_The GUI looks like this -_**  
  
![image](https://user-images.githubusercontent.com/87979848/216362380-d21635af-cedd-4a77-aaeb-33f269dae5ba.png)  
  
**_The Input File Should be in the following format -_**  
  
![image](https://user-images.githubusercontent.com/87979848/216363250-54e5a44c-37c2-403c-aec9-5720bf124ac9.png)  
  
The code first reads all the numbers accross all the sheets and stores it to be used later on. Then, if there is no re-evaluation url, it just works on the given url and 
uses multi-threading to complete the work quicker. Four threads will be initiated for every sheet one after the other and each set will get 25% of the numbers from that 
sheet. Each thread loops over the given numbers and reloads till it gets the data its looking for, for each particular number and stores the scrapped data to a dictionary.  
If a re-evaluation URL is given, then the thread first looks if the number is valid in the given URL, if the data does not show up, it goes to the regular URL and looks 
for the data there and stores it in a dictionary. After the scrapping is done, the code writes the data to an excel file and stores it to the given output location given in 
the GUI.  
  
**_The Output File will look like this -_**  
  
![Screenshot 2023-02-02 205752](https://user-images.githubusercontent.com/87979848/216367778-22752b98-9a5a-4202-9678-0f39924a777b.jpg)  
  
  
