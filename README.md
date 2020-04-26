# AndroidStr
AndroidStr will allow you to get all the strings in your android `strings.xml` files into an excel spreadsheet. And vice versa.

The primary purpose of AndroidStr is to help with creating an excel file of strings for translation. 
Once the translations are ready in the excel file, the same can be used to generate `strings.xml` for each language.

## How to use AndroidStr
AndroidStr is a python script and requires you to have python installed on your computer.

* To check if python is installed, type in the following in terminal/command prompt and hit enter.

    `python --version`

* If you get a valid version number, you are good to go. Else you will have to download python at https://www.python.org/downloads/

## Read all Strings.xml to Excel Sheet

Just provide the path to the `res` folder of your android project and AndroidStr will read the default `strings.xml` as well as the `strings.xml` of all languages you may already have in your project.

1. Download AndroidStr. Open your terminal/command prompt and cd to AndroidStr directory.

2. Once in AndroidStr, type in the following and hit enter.

    `python StringsToSheet.py <path-to-res>`

    \<path-to-res> being the absolute path to the res folder of your android project.

3. An excel file containing all your strings will be created in the AndroidStr directory.
