# split_excel

# Installation  
Create directory to clone `split_excel` project into  
`mkdir ~/split_excel`

Clone `split_excel` into above directory  
`git clone https://github.com/Asciotti/split_excel.git ~/split_excel`

Move into the `split_excel` directory  
`cd ~/split_excel`  

Install dependencies  
`pip install -r requirements.txt` 

### or 
Install via yml file (conda)  
`conda env create -f environment.yml -n desired_name`

# Usage
After navigating to the `src` directory inside of your `split_excel` directory, one can call the file via command line such as:  

`python split_excel.py path/to/original_file.xlsx path/to/output_directory/ split_size`

`split_size` is the number of rows you wish the split the original xlsx file by. If the original xlsx file is not evenly split, the remainder rows are appended to the last subfile. 

# Example
If the working directory looks like:

```
# split_excel/
|--# src/
|    |--# split_excel.py
|
|--# file_to_split.xlsx
```
or 
```
# split_excel
|--# src
|    |--# split_excel.py
|
|--# file_to_split.xlsx
|--# output_dir/
```

If `output_directory` does not exist, it will be created

`python split_excel.py ../file_to_split.xlsx ../output_dir/ 3` will split the `file_to_split.xlsx` by rows of `3` and output the subfiles into `output_dir`
