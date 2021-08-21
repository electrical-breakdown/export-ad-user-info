<#

    .SYNOPSIS
    Retrieves a set of properties for an Active Directory user and exports them to an Excel workbook.

    .DESCRIPTION
    Retrieves a collection of user properties from a specified Active Directory OU and exports that data to a new Excel workbook.

    .INPUTS
    Does not accept any inputs.

    .OUTPUTS
    Outputs a new Excel workbook containing the desired user properties.

#>

# Specify the data to export and the labels to give the corresponding columns in the spreadwriteSheet

$propertiesToRetrieve = "Name", "UserPrincipalName", "StreetAddress", "City", "State", "PostalCode", "Mobile", "Department", "Title", "Manager"
$columnLabels = "Name", "Email Address", "Office Address", "Cell Phone", "Department", "Title", "Reporting Manager"

$targetOU = "OU=EB Users, OU=Electrical Breakdown"
$targetDomain = "DC=electricalbreakdown, DC=com"

$exportFileName = "$($PSScriptRoot)\ADUser_Export.xlsx"  

# Retrieve user objects from AD, including all the properties specified in $propertiesToRetrieve

Write-Verbose "Retrieving user data from Active Directory..."
$adUsers = Get-ADUser -Filter * -SearchBase "$($targetOU), $targetDomain" -Properties $propertiesToRetrieve

    
try {
    
    Write-Verbose "Creating new workbook..."  -Verbose

    # Initialize Excel objects and create a new workbook 
     
    $excelObject = New-Object -ComObject Excel.Application
    $workbook = $excelObject.Workbooks.Add()   
    $writeSheet = $workbook.Worksheets.Item(1)    
   

    # Select the entire first row and give it bold text and a background color

    $firstRow = $writeSheet.Cells(1, 1).EntireRow
    $firstRow.Font.Bold = $true
    $firstRow.Interior.ColorIndex = 15    

     # Loop through $columnLabels and add label to the first cell in each of the columns    

     for($i = 1; $i -le $columnLabels.Count; $i++){

        $writeSheet.Cells(1, $i).Value2 = $columnLabels[$i - 1]
        
    }

}

catch {

    Write-Host "There was a problem creating the workbook." -ForegroundColor Red           
    throw
}

try {

    Write-Verbose "Writing user data to spreadwriteSheet..."  -Verbose   
   
    
    # Loop over all users returned from the call to Get-ADUser and write values to the writeSheet
    # Initialize $row at 2 because the first row contains the column headers  

    $row = 2    

    foreach($adUser in $adUsers){      

        try {
            
            # Extract just the managers name by splitting the string at the = sign
            $adManager = $adUser.Manager.Split(("=",","))[1]    
        }

        catch {
            
            $adManager = ""             
        }

        try {

            # Combine all of the address properties into one string
            $formattedStreetAddress = $adUser.StreetAddress.Trim().Replace("`n", " ")                   
            $adUserAddress = "$formattedStreetAddress $($adUser.City), $($adUser.State) $($adUser.PostalCode)"
        }

        catch {
        
            $adUserAddress = ""            
        }
              
         
        # Collect all the values to write to the worksheet. Be sure to order them to align with the column labels

        $valuesToWrite = $adUser.Name, $adUser.UserPrincipalName, $adUserAddress, $adUser.Mobile, $adUser.Department, $adUser.Title, $adManager

        # Write data to cells

        for($col = 1; $col -le $columnLabels.Count; $col ++){

            $writeSheet.Cells($row, $col).Value2 = $valuesToWrite[$col - 1]

        }                                                                                                                                                                       
                

        #---------Done writing data; increment `$row` and move to next row -------#

        $row += 1

    }  # Close foreach loop


    Write-Verbose "Formatting cells and saving changes to new workbook..."  -Verbose
    

    # Select the columns containing new data and resize them
   
    for($i = 1; $i -le $columnLabels.Count; $i++){

        $writeSheet.Cells(1, $i).EntireColumn.AutoFit() | Out-Null                            
    }

    # Save workbook and close Excel
    
    $workbook.SaveAs($exportFileName)        
    $workbook.Close()
    $excelObject.Quit()

}  # Close try block


catch {

    Write-Host "There was a problem writing data to the workbook. Please ensure the file isn't in use and try again." -ForegroundColor Red
    throw
}


Clear-Host

Write-Host "--------------------------------------------------------------------------"
Write-Host "$($adUsers.Count) users have been exported to: $exportFileName" -ForegroundColor Green
Write-Host "--------------------------------------------------------------------------`n"


Read-Host "Press any key to exit"








