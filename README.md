<<<<<<< HEAD
# stock-analysis

## Overview of Project
This project seeks to provide valuable insights into the refactoring of a VBA Macro, a coded method or function within Excel. 

### Purpose
The purpose of this project is to determine whether or not certain changes allow the program to execute more quickly. The Module 2 Challenge description does a good job of spelling this task out: "to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code[...]. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task."

To this end, we will also take a look at images that show how quickly the code has executed. 

## Results
After refactoring the code, it appears that the refactored code runs more efficiently. To limit the time it takes to run the code to no more than three significant figures, we get the following data points for the original code: 
#### Original Code
*2017 Analysis: 0.715 seconds
*2018 Analysis: 0.707 seconds

#### Refactored Code
*2017 Analysis: 0.168 seconds
*2018 Analysis: 0.109 seconds

### What Gives?
There are several pieces of logic that were reorganized to make the refactored code run nearly 4x more efficiently. The most glaringly obvious one can be found in the for loop structure: In the original code, we loop through all tickers (twelve times) and, for every ticker, we loop through all 3013 rows. That amounts to over 36,000 loops. Furthermore, instead of saving the outputs to an array, we include the output cells *within* the for loop. This may be marginally added computation when compared to the nested for loop described above; but it may also add to the challenge of finding a logical exit to the massive for loop. Notably, the formatting occurs afterward in a different sub routine. 
The refactored code wins the race with its use of a ticker index to replace the higher-level for loop in the original code that manually looped through all twelve tickers. By including if-then logic to increment the ticker once we've reached the end of any given ticker string, such as "AY" to "CSIQ", the computer is able to save its resources: Instead of looping through all 3,000+ rows *every time*, we are able to 'stop' and 'reset' the loop as soon as the ticker increments. By this logic, we reduce the total number of loops tenfold. 

### The Code: Execution Times
One might imagine that the near tenfold reduction in loops between original and refactored code would translate directly to a tenfold reduction in computer processing. Here is where we can boast the fact that all of our formatting was also added to the refactored subroutine, instead of being a separate, post-timer sub() as it was in the original code. Beyond this point, we might also consider the logic reduction and how that translates to reduced computation. This point may be beyond our lowest level of abstraction; we will refrain from analyzing how the if-then statements translate to gate-level logic and processing. However, it's worth noting that some of the added logic in our "longer", refactored loop (e.g., an additional if-then statement) may have added some time to the computation, even if it reduced the total number of loops. For this reason, the refactored code appears to produce a 4x efficiency in computation speed. 

### Other challenges & opportunities
It would be of interest to expound on the above discussion and analyze time stamps at each step of the sub routine. For example, instead of merely starting and stopping the timers and the beginning and end of the code, we might add timers to the beginning and end of any given loop. For the sake of abstraction, we might even run this test on one or two tickers alone, to minimize computational noise that may occur between loops. 

## Summary 
The refactored code performed much more efficiently than the orignal code. Noticeably, the refactored code also included formatting logic in its timing; yet still reduced computation time fourfold. This is likely in large part to our removing the nested for loop that looped all 3,000+ rows for every single ticker. Instaed, we opted to add one piece of logic -- two if-then statements to first confirm if we had arrived at the next ticker, and next to increment the ticker index -- to minimize the number of loops required over all 3,013 rows. 
=======
# stock-analysis
>>>>>>> 586017758c3adcbb232e95d5f6c9566d5d45214d
