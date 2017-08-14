Excel to CSV File Conversion

The application allows the user to convert excel files to CSV files. Each sheet can be separated into different CSV files.
The application also handles splitting 2 tables within a single sheet and saving them in 2 CSV files.


There is a slight hiccup which I couldn't get past regarding this.
If you run the file, you will notice that first column after the empty column (which divides the 2 tables in a single sheet) is being ignored. It is to be noted that only the first column is ignored and the rest of the data is stored into the second CSV file. 
Code wise, there isn't an issue. This seems to be a problem with the ExcelReader that I've used. Instead of reading the data it sends a System.DBNull value which isn't expected. Nevertheless, if the the blank cells are replaced with a specific string "AAA" for example and the code is changed accordingly, the program works like a charm.
