import os
import csv

# Path to collect data from the Resources folder
datafinance_csv = os.path.join('..', 'Resources', 'budget_data.csv') 
# ------------------------------------------------------------------------------
# This simple Python program finds the Max & Min of Profit / Losses of a csv file
# Author: Piruz Alemi, Nov 4th, 2019
#------------------------------------------------------------------------------

total=0
monthCount = 0
profitLoss = 0
maxProfit = 0
maxLoss = 0
maxDate = ""
minDate = ""
holdsDate = ""

beginAmount = 0
endAmount = 0

firstAmount = 0
lastAmount = 0
changeAmount = 0
TotChangeAmount = 0
avgChangeAmount = 0.00

# -------------------------------------------------------------------------------
# Read in the CSV file
with open(datafinance_csv, 'r') as csvfile:

    # Split the data on commas
    csvreader = csv.reader(csvfile, delimiter=',')

    header = next(csvreader)
    print(header)    
#--------------------------------------------------------------------------------
# Loop through the data, at the end print Total, average and max, min values 
# for profit / losses with its corresponding dates... 
# -------------------------------------------------------------------------------
    for row in csvreader:
            # print(row[0],row[1])
            # print(row)
            
            profitLoss = int(row[1])
            holdsDate = str(row[0])
            # print(holdsDate)
            monthCount = monthCount + 1

            if monthCount == 1:
                firstAmount = profitLoss;

            if monthCount > 1:
                endAmount = profitLoss 
                changeAmount = endAmount - beginAmount
                print(holdsDate, changeAmount)
                beginAmount = endAmount;
            else:   
                beginAmount = profitLoss;
               
            # Find the Max/Min Profit/Losses Monthly Changes with its corresponding Max/Min dates
            if changeAmount > maxProfit:
                maxProfit = changeAmount
                maxDate = holdsDate;
            elif changeAmount < maxLoss:
                 maxLoss = changeAmount
                 minDate = holdsDate;
            else:
                pass

            total = total + profitLoss
          
        # End of the For Loop based on indentation

    lastAmount = profitLoss
    totChangeAmount = lastAmount - firstAmount
    avgChangeAmount = (totChangeAmount/(monthCount-1))
   
    # -----------------------------------------------------------------------------
    print("------------------------------------------------------")
    print("-             Financial Analysis                     -")
    print("------------------------------------------------------")
    print("   Total Months: ", monthCount)
    print("   Total: ", "$" + str(total))
    print('   Average Change: ${:,.2f} '.format(avgChangeAmount))
    print("   Greatest Increase in Profits: " + str(maxDate) + " ($" + str(maxProfit) + ")")
    print("   Greatest Decrease in Profits: " + str(minDate) + " ($" + str(maxLoss) + ")")
    print("------------------------------------------------------")

    # --------------------------------------------------------------------------
    # 
    #  Write the results also to an output TXT file, rather than just Terminal
    # 
    # --------------------------------------------------------------------------
    f=open("alemi_outfile.txt","w")
    #---------------------------------------------------------------------------
    result = str("-"*60) 
    L = [result,"\n"] 
    f.writelines(L) 
    # ----------------------------------------------------------------------------
    L = ["                             Financial Analysis                   ","\n"] 
    f.writelines(L) 
    #-----------------------------------------------------------------------------
    result = str("-"*60) 
    L = [result,"\n"] 
    f.writelines(L) 
    #-----------------------------------------------------------------------------
    result = str(monthCount) 
    L = [f"Total Months: ", result,"\n"] 
    f.writelines(L) 
    #---------------------------------------------------------------------------
    result = str(total)
    L = [f"Total: ","$", result,"\n"] 
    f.writelines(L) 
    #--------------------------------------------------------------------------
    # Round the float by 2 decimal points before writing to report
    #-------------------------------------------------------------------------
    x = round(avgChangeAmount,2)
    # print(x)
    result = str(x)
    L =[f"Average Change: " + "$" + result + "\n"]
    f.writelines(L) 
    # -------------------------------------------------------------------------
    result = str(maxProfit) 
    L = [f"Greatest Increase in Profits: " + maxDate + " ($" + result + ")" + "\n"] 
    f.writelines(L) 
    result = str(maxLoss) 
    L = [f"Greatest Decrease in Profits: " + minDate + " ($" + result + ")" + "\n"] 
    f.writelines(L) 
    #---------------------------------------------------------------------------
    result = str("-"*60) 
    L = [result,"\n"] 
    f.writelines(L) 
    # ----------------------------------------------------------------------------
    f.close()
    