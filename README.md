# The VBA of Wall Street

## Background

I initially solved this with a very basic solution wherein I kept recalculating the stock change and stock volume and inserting that result into a new column. This let me watch the script run but was incredibly slow. 

Once I felt comfortable with the solution, I rewrote it by moving processes and calculations outside of the for-loop, which meant rewriting a few thing and switching to ranges instead of cells. 

For the Master Results, I used a second for-j loop. This isn't the most efficient but I wanted to try the difficult solution. It's a rewrite of the summary table I had prior using ranges instead of cells. Unfortunatelly, you have to run the code twice to see the Master Results. My hope is to find a better way to initiate the checks within the i-loop. 

### About This Repo

I've included the test script folder and the final script folders. There are also screenshots. 

