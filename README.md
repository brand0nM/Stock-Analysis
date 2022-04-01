# Stock-Analysis
## Project Overview
Previously I created a workbook that can analyze a dozen stocks using VBA; Now I wants to expand this dataset to include the entire stock market over the last couple years.

### Purpose
To do this, we'll refactor the previous solution to loop over our data, collecting unique ticker names, then determine their volume/annual growth rate. Finally, we’ll determine whether refactoring our code was successful in making the VBA script run faster; a secondary goal of this refactor is to avoid hard-coding every element to our Tickers() array- this would take exponentially more time with thousands of ticker names. 

## Analysis

### Analysis of 2017

##### Original
<img width="870" alt="Their_Code,2017" src="https://user-images.githubusercontent.com/79609464/161325218-cc353230-52d9-4f30-82c9-96a4d9cdf88a.png"> 

##### Refactored
<img width="900" alt="My_Code,2017" src="https://user-images.githubusercontent.com/79609464/161325213-37ac4302-e97d-4a63-90a1-4041ebc8902f.png">

### Analysis of 2018

##### Original 
<img width="870" alt="Their_Code,2018" src="https://user-images.githubusercontent.com/79609464/161325285-ae38ef3c-58db-4e2b-9b0c-603ad5564fd0.png"> 

##### Refactored
<img width="840" alt="My_Code,2018" src="https://user-images.githubusercontent.com/79609464/161325283-0f32a60e-d7a1-4dee-8db4-6ab707cf6143.png">

## Results
Though my refactored code is 5-6 times slower than our previous code, this is primarily due to that fact that we aren't hard coding ticker names to the Tickers() array; Instead our code loops through the data twice, once to find unique ticker names adding them to our array, and again to find each tickers volume and growth rate over the past year. In the long run this will save countless hours, as we no longer have to manually add thousand of tickers to our array. 

### Challenges and Difficulties Encountered
Data types in VBA are not as intuive as a python; for example, adding an element to an array is much more involved than python. In python we simply declare an empty list or dictionary and append unique values to it, but in VBA we must declare the dimensions and whether or not we want to preserve those dimensions. 
