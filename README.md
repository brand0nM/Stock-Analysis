# Stock-Analysis
## Project Overview
We've already created a workbook that can analyze a dozen stocks or so stocks using VBA; Now we wants to expand this dataset to include the entire stock market over the last few years.

### Purpose
In this challenge we'll refactor our previous solution to loop through all the data one time in order to collect the same information that you did in this module. Then, we’ll determine whether refactoring our code successfully made the VBA script run faster. When refactoring code, we want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 

## Analysis and Challenges

### Analysis of 2017

<img width="900" alt="My_Code,2017" src="https://user-images.githubusercontent.com/79609464/161325213-37ac4302-e97d-4a63-90a1-4041ebc8902f.png">
<img width="870" alt="Their_Code,2017" src="https://user-images.githubusercontent.com/79609464/161325218-cc353230-52d9-4f30-82c9-96a4d9cdf88a.png">


### Analysis of 2018

<img width="840" alt="My_Code,2018" src="https://user-images.githubusercontent.com/79609464/161325283-0f32a60e-d7a1-4dee-8db4-6ab707cf6143.png">
<img width="870" alt="Their_Code,2018" src="https://user-images.githubusercontent.com/79609464/161325285-ae38ef3c-58db-4e2b-9b0c-603ad5564fd0.png">


### Challenges and Difficulties Encountered
Data types in VBA are not as intuive as a python; for example, simply adding an element to an array is much more involved than python. In python we can simply declare an empty list or dictionary and append unique values to it, but in VBA we must declare the dimensions and whether or not we want to preserve those dimensions. Both can do similar things, but python would have taken 3 line of code, where this took 10.

## Results
Though my newly refactored code is marginally slower than our previous code, this is soley due to that fact that we aren't hard coding ticker names to the Tickers() array; Instead our code loops through the data twice, once to find unique ticker names, adding them to our array, and again to find each tickers volume and growth rate over the past year. In the long run this will save countless hours, as we no longer have to manually add tens of thousand of tickers to our array.
