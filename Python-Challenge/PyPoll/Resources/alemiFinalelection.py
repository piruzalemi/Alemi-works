import os
import csv

# Path to collect data from the Resources folder
datafinance_csv = os.path.join('..', 'Resources', 'election_data.csv') 
# ----------------------------------------------------------------------------
#                                BONUS !!!!!
# The dataset is composed of three columns: Voter ID, County, and Candidate. 
# This Python script analyzes the votes and calculates each of the following:
#
#         1. The total number of votes cast
#         2. A complete list of candidates who received votes
#         3. The percentage of votes each candidate won
#         4. The total number of votes each candidate won
#         5.  The winner of the election based on popular vote.
#         6. Exports the results to output file
#                                          Author: Piruz Alemi, Nov 4th, 2019
#----------------------------------------------------------------------------

voterID = ""
county = ""
candidate = ''
totalVotesCast = 0
listOfCandNames =[]
percentCandVotes = 0.000
totalCandVotes = 0
theElectionWinner = ""
i = 0
maxVote = 0
maxName = ""
# Create a Dictionary where its key is Candidate and its values are Voter ID counts
dictCandid ={
    }


# -------------------------------------------------
# Read in the CSV file
with open(datafinance_csv, 'r') as csvfile:

    # Split the data on commas
    csvreader = csv.reader(csvfile, delimiter=',')

    header = next(csvreader)
    print(header)
    
    
    #--------------------------------------------------------------------
    # Loop through the data, at the end create the report
    # -------------------------------------------------------------------
    # Dictionary key candidate name: str value vote_count int
    for row in csvreader:
            # print(row[0],row[1])
            # print(row)
            
            voterID = int(row[0])
            candidate = str(row[2])
            # print(holdsDate)
            totalVotesCast = totalVotesCast + 1

            if(candidate == None or candidate == ''):
                continue;
            if(candidate not in dictCandid):
                dictCandid[candidate] = 1;
            else: 
                dictCandid[candidate] = dictCandid[candidate] + 1
           
            i = i + 1
      
            #if i == 50:
            #   break             
    
                    
        # End of the For Loop based on indentation
    print(dictCandid)
    for candidate in dictCandid:
        #print(candidate, '-->', dictCandid[candidate])
        if dictCandid[candidate] > maxVote:
            maxVote = dictCandid[candidate]
            maxName = candidate;
            
    #print('winner is!!!!!!!!: ' + str(maxVote), str(maxName))
    # --------------------------------------------------------------------------
    #                         Print the Results
    # --------------------------------------------------------------------------
    print("-"*60)
    print("-Election Results                       -")
    print("-"*60)
    print("Total Votes: ", totalVotesCast)
    for candidate in dictCandid:
        percentCandVotes = round(dictCandid[candidate]/totalVotesCast,3)
        x=percentCandVotes
        print(candidate,':           ', "{:.0%}".format(percentCandVotes), '(',dictCandid[candidate],')')

    print("-"*60)
    print(f'Winner: {maxName}')
    print("-"*60)
    # --------------------------------------------------------------------------
    #                 Write these results also to an output CSV file
    #                 In the PyBank the output was written to a txt file
    # --------------------------------------------------------------------------
    # output_path = os.path.join("..", "output", "new.csv")
    output_path = os.path.join("..", "Resources", "alemi_candidate_outfile.csv")

    # Open the file using "write" mode. Specify the variable to hold the contents
    with open(output_path, 'w', newline='') as csvfile:
    # Initialize csv.writer
        csvwriter = csv.writer(csvfile, delimiter=',')
    # Write the first row (column headers) 
        csvwriter.writerow(['Total Votes', 'Candidate Name', ' Candidate Votes','Percent Candidate Vote','Winner'])
        #---------------------------------------------------------------------------------------------------------
        # Write loop to report on candidates, from the Candidate Dictionary
        # Note!!!!!!!!!!  The for Loop must be indented for the Open File, otherwise we will have an I/O error
        #---------------------------------------------------------------------------------------------------------
        for candidate in dictCandid:
            percentCandVotes = round(dictCandid[candidate]/totalVotesCast,3)
            z="{:.0%}".format(percentCandVotes)
            
            #csvwriter.writerow([candidate,z])
            csvwriter.writerow([totalVotesCast, candidate,dictCandid[candidate], z, maxName,'\n'])

