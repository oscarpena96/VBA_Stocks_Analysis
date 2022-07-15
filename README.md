
Introduction
In this project we are going to analyze the Stock Market with VBA creating a first draft to understand the way VBA works and how to utilize its tools. After that we are going to refactor our first code to make it better measuring the time taking that parameter as reference for better coding.
The code will be able to calculate the stock price change trough the year, the difference in percentage from the beginning of the year to the end and will calculate the total volume traded for the year.
Background
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.
In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.
Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

Results
For my results I got 2 excels, green_books and VBA_Challenge, the first one was for me to be able to run the code so this one was for practicing, with this code I was able to run it in 0.68 (fig1 and 2) seconds for both years (2017 and 2018). My code was slow do to 2 for loops (fig. 3) in this case thanks to the loops my code is running a total of 33,143 times.

For my second code, I did some changes to get a better performance, thanks to this the code was able to run in 0.109 seconds and 0.101 seconds for the years 2017 and 2018 respectively, this was an improvement of 0.5 seconds for both years.


In the code there is only one for loop so it´s only going to the data one time, that’s a total of 3013 cells so depending on your machine it is going to be able to run even faster than in mine


Refactoring
Advantages 
It´s easier to see the mistakes that you commit
It´s easier to read the code and explain it to somebody else
Faster to run and that´s better for big data 
Simplify code
Disadvantages
You need a lot of practice to be able to perfect your code
You need to know what you are doing 
Imprecise factoring can bring bugs or errors

The way I see it VBA it´s not use like other programming languages, basically it´s use for small codes to automatize some procedure so refactoring it´s good to make it easier to read but because its not a lot of work that you normally do, I don´t see the need to refactor unless you are going to give it to some company or somebody else
