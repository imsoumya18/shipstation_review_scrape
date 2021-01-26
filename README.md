[![forthebadge made-with-python](http://ForTheBadge.com/images/badges/made-with-python.svg)](https://www.python.org/)

# Shipstation Review & Scrape
An automation script to automate browser that will review invoices from given excel sheet

***Steps:***
1. Keep the file you want to modify in the same folder as the main.py is in.
   For example here it is **2032442-FDBackup.xlsx** included. Replace that with your original file.
2. After running the script it will first ask for date. Enter it in **dd-mm-yyyy** format (e.g. 12-20-2020 for 20th December, 2020).
3. Next, it will ask for invoice no. Enter only the no without any sign.
   For example, **2037471**. Don't enter any sign like **#** or anything.
4. Then, it will ask for the name of the file you want to modify. Just type name only with extension (e.g. **2032442-FDBackup.xlsx**) if it's in the same folder. Otherwise give the full path.
5. Now, a chrome page will open up and it will open the login page. Enter your Shipstation **username** and **password** and click **Login**.
6. Now, it will be automatically logged in in a few seconds. In this time come back to your code. There it will say to click ENTER after **page has completely loaded**. Remember, **the loading wheel must have vanished and the list of all records must be loaded and then only you have to hit ENTER**.
7. Now, it will do the rest automatically. You will get an excel file named in the format **GLOBEGISTICS 12-20-2020 - INVOICE 2037471** for the above given example.

***Precautions:***
8. Before running the script login to your Shipstation account and make sure **Show Sidebar** is enabled. Otherwise code will fail.
9. Also, make sure **Store** column is enabled.

In case of any problem mail me: <a href="mailto:soumyadeep184@gmail.com">soumyadeep184@gmail.com</a> 





![forthebadge made-by-soumya](https://img.shields.io/badge/CREATED_BY-SOUMYA-blue)



I got this project as a freelance job from **Nicolas Troplent** on **Upwork**. Thanks to him❤. I shall be glad to work with him again.

![forthebadge thanks-to-nico](https://img.shields.io/badge/THANKS_TO-NICO-brightgreen)
