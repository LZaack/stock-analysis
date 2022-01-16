# Stock Analysis with VBA

## Overview of Project
Steve needed to analyze stock data to help his parents to make the best investment decisions. He could have performed the data analysis manually and using Excel formulas, but to automatize the analysis Steve prefer to ask us to create a code with VBA to do it faster and decrease errors.

Steve was very glad about our job, but then he needed to expand the data set to include the entire stock market over the last few years and give his parents a complete overview of the market, so he asked us to refactor the code to accomplish the following goals:
  - Process thousands of stock information.
  - Make it more efficient, so it runs faster, takes fewer steps, and uses less memory.

## Results
The code from outside in is now as follows:

`1. First Step` loop over all the rows in the spreadsheet and increase the volume of the running ticket. We use the next `for` loop:

![Loop_Increase_Volume](https://user-images.githubusercontent.com/95454286/149607821-a6c23e4a-05ba-43d9-89fa-85d44ac08f4c.png)

`2. Second Step` verifies if the current row is in the first row selected by the tickerIndex: We use the next `condition` code:

![Check_First_Row](https://user-images.githubusercontent.com/95454286/149608286-f689205c-3a7d-4978-a545-8d19fac67962.png)

`3. Third Step` verifies if the current row is in the last row selected by the tickerIndex: We use the next `if` code:

![Ckeck_Last_Row](https://user-images.githubusercontent.com/95454286/149608285-5d718e14-8d6f-45fa-a26d-11ff17a5d789.png)

`4. Fourth Step` checks if the tickerIndex is still the same, if it is not, it increases the tickerIndex to past to the next value of the tickerIndex: We use the next `if` code:

![Increase_TickerIndex](https://user-images.githubusercontent.com/95454286/149608300-f9fa855c-1acf-4379-9680-b47569e6f87b.png)

`5 Fifth Step` print the Ticker, Total Daily Value, and Return, in the All Stock Analysis worksheet. To do so we use the following `loop` code:

![Loop_Output](https://user-images.githubusercontent.com/95454286/149608280-86443b54-d95f-46e2-90fa-f356a38765b7.png)

The previous code ran in 1.171875 seconds, and the refactored code run in 0.3007813 for the year 2017, as it can be shown:

![Code_Ran_Result_2017](https://user-images.githubusercontent.com/95454286/149608590-8f890123-05a7-4437-85b4-9f553a944f87.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/95454286/149608592-13e71cd3-3c90-4c26-a54e-31e209c1787e.png)

In the year 2018 the previous code ran in 1.140625 seconds, but in the refactored code it only takes 0.2480469 seconds, how in the next images is going to be proved:

![Code_Ran_Result_2018](https://user-images.githubusercontent.com/95454286/149608630-9a3348a6-ed31-4bdf-babc-c892435765e4.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/95454286/149608631-558611d7-9b61-4467-a2d4-068ddd6a0058.png)

## Summary

We achieve the same result with both codes, but with the refactored one we accomplished some Steve´s Goals, like making it more efficient, so it can run faster and use less memory.

### 1. What are the advantages or disadvantages of refactoring code?

*Advantages:* It can process more information in less time. It could be used as a template.

*Disadvantages:* Takes a lot of time to refactor and make it run properly without errors. Probably for a huge project the time invested is worth it.

### 2. How do these pros and cons apply to refactoring the original VBA script?

*Cons:* The main problem is to structure how you are going to use the variables, the values that you are going to assign them, and define which step is carried out first from the inside out.

*Pros:* Once you are clear with the steps and the variable’s values, the code could be used in many other similar projects, with a few changes.
