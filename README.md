This project started from a base dataset that was cleaned before pulling the applications. However, the code will work even if there is nothing in a csv that//
you want to populate with the variables in the code. 

The "Pulling Applications" file is the code that was used to populate the analysis csv. Any further analysis can be done on this csv, or even add more applications//
to it. 

It is important to note that the applications do have to be downloaded from TCAC and saved into a folder(s) to be able to pull the data of interest. 

The "Analysis Portion" is where further analysis will be done based on our goals. More will be added to this file later

After making any changes and having to rerun the code, the output file will not update. 
- You will need to change the format of the excel output to make it viewable by freezing cells, adding filters, and adding this under to get the min, max, and mean on the bottom lines
- =SUBTOTAL(101, top cell:bottom cell) for averages
- =SUBTOTAL(104, top cell:bottom cell) for max
- =SUBTOTAL(105, top cell:bottom cell) for min
