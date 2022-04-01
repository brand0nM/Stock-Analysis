# Stock-Analysis
## Project Overview
Previously I created a workbook that can analyze a dozen stocks using VBA; Now I wants to expand this dataset to include the entire stock market over the last couple years.

### Purpose
To do this, we'll refactor the previous solution to loop over our data, collecting unique ticker names, then determine their volume/annual growth rate. Finally, we’ll determine whether refactoring our code was successful in making the VBA script run faster.

## Analysis and Challenges

### Analysis of 2017

<img width="900" alt="My_Code,2017" src="https://user-images.githubusercontent.com/79609464/161325213-37ac4302-e97d-4a63-90a1-4041ebc8902f.png"> Refactored <br />
<br />
<img width="870" alt="Their_Code,2017" src="https://user-images.githubusercontent.com/79609464/161325218-cc353230-52d9-4f30-82c9-96a4d9cdf88a.png"> Original <br />


### Analysis of 2018

<img width="840" alt="My_Code,2018" src="https://user-images.githubusercontent.com/79609464/161325283-0f32a60e-d7a1-4dee-8db4-6ab707cf6143.png"> Refactored <br />
<br />
<img width="870" alt="Their_Code,2018" src="https://user-images.githubusercontent.com/79609464/161325285-ae38ef3c-58db-4e2b-9b0c-603ad5564fd0.png"> Original <br />

We can observe from the 
### Challenges and Difficulties Encountered
Data types in VBA are not as intuive as a python; for example, simply adding an element to an array is much more involved than python. In python we can simply declare an empty list or dictionary and append unique values to it, but in VBA we must declare the dimensions and whether or not we want to preserve those dimensions. Both can do similar things, but python would have taken 3 line of code, where this took 10. 

## Results
Though my newly refactored code is marginally slower than our previous code, this is soley due to that fact that we aren't hard coding ticker names to the Tickers() array; Instead our code loops through the data twice, once to find unique ticker names, adding them to our array, and again to find each tickers volume and growth rate over the past year. In the long run this will save countless hours, as we no longer have to manually add tens of thousand of tickers to our array. 
