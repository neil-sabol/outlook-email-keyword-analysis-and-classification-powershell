# This is a PowerShell script that provides a quick, systematic way to get a baseline of what you email about most often.
# It also analyzes how email messages you send (containing specific keywords) are related to one another. It allows you
# to analyze sent email messages over time intervals in Outlook, produce a breakdown of user-defined keywords, and
# display the results in an easy to read format. The idea is similar to a word cloud or tag cloud, but you define the
# search and matching criteria (key words or tags).
# Neil Sabol (neil.sabol@gmail.com)
#
# Based on and inspired by Ed Wilson's (Microsoft Scripting Guy) work:
# https://blogs.technet.microsoft.com/heyscriptingguy/2011/05/26/use-powershell-to-data-mine-your-outlook-inbox/

################################################################################################################
# BEGIN CONFIGURATION - ADD YOUR KEYWORD LIST (GROUPS) HERE
################################################################################################################

# Define projects, tasks, topics, customers, accounts, people, etc. and the keywords associated with them - the script
# uses this array to classify your sent messages (based on keyword match). See the example below and create your
# $keywordList array accordingly. Add as many lines as needed. Keep in mind the first element of each line is the
# LABEL for that group of keywords (and IS NOT included when matching but appears in the "Breakdown by Keyword (relative)"
# section of the script's output). The keywords themselves must be listed after the label (first element) - these ARE
# matched against the content of your sent items.
#
#         $keywordList =  ("Label 1","keyword 1.1","keyword 1.2","keyword 1.3"),
#                         ("Label 2","keyword 2.1")
#

$keywordList =  ("Sycamore Property","2897 sycamore st","sycamore plaza","98123135135"),
                ("Elm Property","elm","elm center","elm apartment"),
                ("Downtown Property","1 civic plaza","downtown office","9889231"),
                ("Santa Fe Property","santa fe","plaza")

################################################################################################################
# END CONFIGURATION
################################################################################################################

# Ensure Outlook is running - technically, Outlook must also be configured and a user logged in
$outlookProcess = Get-Process outlook -ErrorAction SilentlyContinue
if (!$outlookProcess) {
    write-host " "
    write-host "It does not appear that Outlook is running. Please configure your default email profile then"
    write-host "launch and log into Outlook. This script will now terminate."
    pause
    exit
}

# Inform user that parsing of Outlook data has begun (and may take a few moments)
write-host " "
write-host "Please wait, staging Outlook data. This may take 5 minutes or more (depending on your settings and email volume)..."
write-host " "

# VERY minor tweak to Ed Wilson's (Microsoft Scripting Guy) function to extract Outlook Inbox objects (extract Sent Items instead)
# Additional folder names available here: https://msdn.microsoft.com/en-us/library/office/ff861868.aspx
# MailItem properties available here: https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-object-outlook
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
$outlook = new-object -comobject outlook.application 
$namespace = $outlook.GetNameSpace("MAPI") 
$folder = $namespace.getDefaultFolder($olFolders::olFolderSentMail) 
$sentItemsCache = $folder.items | Select-Object -Property Subject, ReceivedTime, Body, To

# Since it takes a while to load Outlook data, provide an opportunity to analyze multiple date ranges in a single session
# This is accomplished with the while loop (and a prompt at the end of each iteration to analyze another date range).
$continueRun="y"
while ($continueRun -eq "y") {

    # Initialize variables
    $currentCount = 0
    $totalEmailsParsed = 0
    $totalMatches = 0
    $totalKeywords = 0

    # Initialize the count array - add a counter (starting at 0) for each keyword group provided in the configuration
    # Assume that keys in this array line up with keys in the keyword array since the latter cannot change during
    # execution
    $keywordCount = @()
    ForEach ($group in $keywordList) {
        $keywordCount+=0
    }

    # Get date range from user
    write-host " "
    $startDate=Read-Host -Prompt "Enter START Date (mm/dd/yy)"
    $endDate=Read-Host -Prompt "Enter END Date (mm/dd/yy)"

    # Inform the user of a potential delay in parsing data (generally this is quick, but just in case)
    write-host " "
    write-host "Please wait, analyzing data for specified dates..."

    # Process sent items, counting the frequency of keywords
    # Pull messages objects out of the "cache" created above, based on the date range entered by the user
    $sentItemsCache | where-object { $_.ReceivedTime -ge $startDate } | where-object { $_.ReceivedTime -le $endDate } | ForEach-Object {
        $totalEmailsParsed++
        $emailSubject = $_.Subject
        $emailBody = $_.Body
        $emailRecipient = $_.To
        $matchFound = "N"

        # Loop through each keyword group (outer loop) and each keyword in that group (inner loop) and check to see if any of the
        # specified email fields contain the keyword. Note, since the first element of the keyword array is the label, skip it - 
        # the $j=1 instead of $j=0 bit
        For ($i=0; $i -lt $keywordList.Length; $i++) {
            # Get the count for the current keyword from the keyword count array
            $currentCount = $keywordCount[$i]
            For ($j=1; $j -lt $keywordList[$i].Length; $j++) {
                # Prepend and append the wildcard character for -like matching - do so for each keyword in the keyword group (besides)
                # the label (at position 0)
                $currentKeyword = "*" + $keywordList[$i][$j] + "*"
                if ($emailSubject -like $currentKeyword -Or $emailBody -like $currentKeyword -Or $emailRecipient -like $currentKeyword) {
                    # Check the keyword against email fields - if matched, update the count
                    $matchFound = "Y"
                    $keywordCount[$i] = $currentCount+1
                }
            }
        }
        # If any keywords in the group matched this message, update the count of total matches
        if ($matchFound -eq "Y") {
            $totalMatches++
        }
    }

    # Tally total matches across all keyword groups to normalize the results (and calculate %)
    For ($i=0; $i -lt $keywordList.Length; $i++) {
        $currentCount = $keywordCount[$i]
        $totalKeywords = $totalKeywords + $currentCount
    }

    # If any messages were parsed, estimate % of emails matched: the number of emails matched divided by the total
    # emails parsed
    if ($totalEmailsParsed -ne 0) {
         $percentMatches = [math]::Round(100 *($totalMatches/$totalEmailsParsed))
    
        # Print summary by keyword
        write-host " "
        write-host "Breakdown by Keyword (relative)"
        write-host "-------------------------------"
        # Loop through the keywords array and keyword match count array, calculate the %, and if greater than 0,
        # print the result
        For ($i=0; $i -lt $keywordList.Length; $i++) {
            $currentKeyword = $keywordList[$i][0]
            $currentCount = $keywordCount[$i]
            $currentKeywordPercentage = [math]::Round(100 * ($currentCount/$totalKeywords))
            if ($currentKeywordPercentage -ne 0) {
                write-host "$currentKeywordPercentage% $currentKeyword"
            }
        }

        # Provide rough quality data - useful for tuning and refining keywords
        # The higher the % match, the better your set of keywords
        write-host " "
        write-host "Messages Analyzed: $totalEmailsParsed"
        write-host "Items matching keywords (estimate): $percentMatches%"
        write-host " "
        write-host " "
    
    # Inform user that no messages matched the dates provided and start over 
    } else {
        write-host " "
        write-host "No messages found in this date range"
        write-host " "
        write-host " "
    }
    
    # Provide an opportunity for additional date range analysis
    $continueRun=Read-Host -Prompt "Analysis complete. Would you like to enter another date range (y/n)?"
}

# Display a friendly reminder to copy/paste analysis results before the script window closes
write-host " "
write-host " "
write-host "Script complete - be sure to copy/paste the analysis above (as needed)"
write-host " "
pause
