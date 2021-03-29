# stock-analysis
These macros are used by a friend, Steve, wishing to analyze the best stocks for his parents to invest in for the years 2017 and 2018. The parents were particularly interested in green energy stocks, especially the stock labeled DAQO("DQ"). I wanted to refactor the code that was used previously in order to improve the runtime and readability.

## Results

### Stock performance
Generally speaking the return is often considered more than the volume of trades, but both are included. After running the macro, DQ seemed to do relatively well in 2017 but poorly in 2018, so adding some diversity to Steve's parents' stock portfolio would do them some good. Possibly adding at least adding the stock ticker ENPH would help.

`tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value`
`Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1`

[2017 analysis](https://github.com/MServ/stock-analysis/blob/main/Resources/VBA_Analysis_2017.png)
[2018 analysis](https://github.com/MServ/stock-analysis/blob/main/Resources/VBA_Analysis_2018.png)

### Execution time improvement
The execution time from the original to refactored code improved from taking over 0.9 seconds to less than 0.2 seconds.

[2017 original](https://github.com/MServ/stock-analysis/blob/main/Resources/VBA_Original_2017.png)
[2018 original](https://github.com/MServ/stock-analysis/blob/main/Resources/VBA_Original_2018.png)
[2017 refactored](https://github.com/MServ/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
[2018 refactored](https://github.com/MServ/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and disadvantages of refactoring code
There are a couple of main reasons to refactor code. These include improving the speed which the scripts will run, which may be of greater importance for larger blocks of code, as well as making the code more easily readable for anyone who needs to look over the code.

### Refactoring for this script
For this particular set of analysis, there was no great need for refactoring in terms of execution time of the script. It took much longer to rewrite the code than would likely be saved by shortening the run speed. However it did improve the clarity and flow of the script itself.