# extract_excel_csv_data_recursively
A very basic GUI app for extracting data from multiple source files into a single csv file

## Function
[Explanation of the underlying function](https://mgw.dumatics.com/read-excel-csv-recursive/)

## Usage
The script can directly be copied to a Jupyter cell or can be run from terminal. Following command should ensure all dependencies are installed:

```python
py -m pip install pandas, numpy, pysimplegui
```

Some things the GUI takes care of are: 

* Allows selection of columns to be extracted from a sample `.csv` file
* Allows user to specify which of the selected columns should be parsed as `date`
* Gives a date based filename to `output`
* Shows colour coded log for which files were read in green and which were ignored in red background.

## Screenshots

### Empty Form
![image](https://user-images.githubusercontent.com/1966557/232877297-2d3a2914-8a7f-4f20-bba1-af8d0e92c039.png)

### Filled Form
![image](https://user-images.githubusercontent.com/1966557/232879216-22e3b81c-10d0-449f-80d1-30725c24b5ff.png)

### Displays extracted output
![image](https://user-images.githubusercontent.com/1966557/232881326-f2ecdd14-8bf6-4f7b-bcd4-3e5de4f361a2.png)

### Displays filename of the output and location where it is saved
![image](https://user-images.githubusercontent.com/1966557/232881685-6a0eafd9-a5be-464b-ae0c-5814b101fc02.png)

### Colour coded log
![image](https://user-images.githubusercontent.com/1966557/232882657-dd1e4e69-26a3-4849-8e94-9811cedab6d7.png)

