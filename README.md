# stock-analysis

# Overview of Project: Explain the purpose of this analysis.
The purpose of this Project is to refactor the VBA code that was written in Module 2. The refactoring process will make the code more efficient in accomodating a larger dataset. The refactoring process is a common part of programming, as the first draft of the code is not likely the final draft. 

# Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

Original 2017 Time: 1.42 Seconds
*Image:     ![image](https://user-images.githubusercontent.com/8634824/121790519-cdef8180-cba5-11eb-80c8-ea5bcad57ffc.png)

Refactored 2017 Time: .5 Seconds 
*Image:    ![image](https://user-images.githubusercontent.com/8634824/121790555-36d6f980-cba6-11eb-838e-654e7a907415.png)  
  
  
  

Original 2018 Time: 1.33 Seconds
*Image:     ![image](https://user-images.githubusercontent.com/8634824/121790562-4eae7d80-cba6-11eb-843a-1472bfc8ae84.png)

Refactored 2018 Time: .6 Seconds
*Image:      ![image](https://user-images.githubusercontent.com/8634824/121790581-71d92d00-cba6-11eb-960e-bc20b77dfd34.png)


**2017 V. 2018 Comparison**

The 2 years were quite opposites. In 2017, All stocks had risen considerably from the previous year, with the most successful Stock being DQ with a 199% return. Only the TERP stock had fallen from the previous year (-7.2%) 
*Image:     ![image](https://user-images.githubusercontent.com/8634824/121790950-540dc700-cbaa-11eb-924e-934e956d15fc.png)


In 2018, All stocks fell considerably from the previous year (2017), with DQ taking the biggest fall at (-62.6%). The only stocks that had a positive return were ENPH (+81.9%) and RUN ( +84%)
*Image:     ![image](https://user-images.githubusercontent.com/8634824/121790955-67209700-cbaa-11eb-8b88-ffbad066ae98.png)

**Code Comparison**
The fundamental difference between the Original and Refactored code seems to be this portion of the Codes:

#Original Code
'5) Loop through rows in the data.

Worksheets(yearValue).Activate
  For j = 2 to RowCount
  
'5A) Find total volume for the current kicker. 
If Cells(J,1).Value = ticker Then
  totalVolume = totalVolume + Cells(j,8).Value
  End If
  
  
#Refactored Code
''2B Loop over all the rows in the spreadsheet.
Worksheets(yearValue).Activate
  For i = 2 to RowCount
   '3a) Increase volume for current ticker
      tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i,8).Value
 



# Summary: In a summary statement, address the following questions.
1. What are the advantages or disadvantages of refactoring code?

***Advantages***
Efficiency in Run Time
Cleaning up the Code makes it easier to understand for the creator and any additional users.

***Disadvantages***
Accidentally effect other portions of the original code
Perhaps Refactoring a section of the code makes it better at the moment, but if you make future edits, the original code would've been better to use.  


2. How do these pros and cons apply to refactoring the original VBA script
The advantages were shown in a significant drop in Run Time , as well as Simplicity in the code. 
