
def colOneSum(sheet):
    # Initialize the sum
    sum = 0
    # Loop through the rows
    for row in sheet:
        # Add the first element to the sum
        sum += float(row[0])
    # Return the sum
    return sum

# Create a list of lists from the file
def parseSheet(fileName):
    # Open the file
    sheet = []
    with open(fileName, 'r') as f:
        # Read the file
        i = 0
        # Read the first line
        for line in f:
            # Skip the first line
            if i == 0:
                i += 1
                continue
            # Split the line into a list
            # Remove the newline character
            # Add the list to the sheet
            sheet.append(line.split(','))

    # Return the list
    return sheet

# Print the sum of the first column
print(colOneSum(parseSheet('covid.csv')))