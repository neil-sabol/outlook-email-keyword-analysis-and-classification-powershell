# Outlook email (sent item) keyword analysis and classification with PowerShell
Ever wonder what projects, customers, or accounts you email about most frequently? This script allows
you to analyze sent email messages over specific time periods in Outlook, produce a breakdown of keywords
you define, and display the results in an easy to read format. Sort of like a word cloud, but you define
the search and matching criteria (key words). Quick way to get a baseline of what topics you email about
most frequently and how the email messages you send for each key word group (label) relate to one another.

Note, this type of analysis assumes you diligently include the keywords you plan to search for as you compose
and send email messages.

Lots of room for improvement here - feel free to fork and pull if you have suggestions or ideas.

## Requirements
* **Outlook client** (not web access) installed and configured with a default email profile (see [Outlook Email Setup](https://support.office.com/en-us/article/Outlook-email-setup-6e27792a-9267-4aa4-8bb6-c84ef146101b)
* Outlook client running (and logged in) at time script is run
* Default email profile set to "Cached Mode" with the amount of mail stored offline set to the amount of mail you wish to process with this script (Control Panel -> Mail):

![Outlook client cache (offline) setting](https://www.binaryinnards.com/images/outlook-cached-exchange-mode-date-range.png)

* PowerShell V2 or higher
* PowerShell execution policy that permits this script to run

## Usage
1. Download this script (Analyze-Outlook-Sent-Item-Keywords.ps1)
2. Open/edit (do not run) the script and setup your $keywordList accordingly - see *Examples* below
3. Run the script - when prompted, enter a start date and end date and review the resulting data

## Sample Run (output)
As a test, I defined some sample categories and ran this script against my Sent Items at work - the following was the result.
```
  ___        _   _             _      _____                 _ _
 / _ \ _   _| |_| | ___   ___ | | __ | ____|_ __ ___   __ _(_| |
| | | | | | | __| |/ _ \ / _ \| |/ / |  _| | '_  _ \ / _ | | |
| |_| | |_| | |_| | (_) | (_) |   <  | |___| | | | | | (_| | | |
 \___/ \__,_|\__|_|\___/ \___/|_|\_\ |_____|_| |_| |_|\__,_|_|_|
   / \   _ __   __ _| |_   _ _______ _ __
  / _ \ | '_ \ / _ | | | | |_  / _ | '__|
 / ___ \| | | | (_| | | |_| |/ |  __| |
/_/   \_|_| |_|\__,_|_|\__, /___\___|_|

Please wait, staging Outlook data. This may take 5 minutes or more (depending on your settings and email volume)...


Enter START Date (mm/dd/yy): 6/19/17
Enter END Date (mm/dd/yy): 6/21/17

Please wait, analyzing data for specified dates...

Breakdown by Keyword (relative)
-------------------------------
6% Content Management System
8% Ticketing System
6% Wiki
21% Misc Apps
19% Knowledge Base
40% Websites and Hosting

Messages Analyzed: 69
Items matching keywords (estimate): 74%


Analysis complete. Would you like to enter another date range (y/n)?:
```

## Examples (for $keywordList)

### Using names (people)
May apply if you want to see who you send email to or email others about
```
$keywordList =  ("Bob Bobson","bob bobson","robert bobson","bob"),
                ("Jim Jimson","james jimson"),
                ("Jane Janeson","jenny","jane janeson"),
                ("Jill Jillson","jill jillson","jill")
```
Note, the first element of each line (i.e. "Bob Bobson", "Jim Jimson", "Jane Janeson", etc.) is the label for that group of keywords - it is not included when searching your sent messages. The keywords (i.e. "bob bobson", "robert bobson", "jenny", etc.) follow the label - these ARE matched against the content of your sent items.

### Using projects
May apply if you work on several projects, tasks, and activities
```
$keywordList =  ("Bakery Website","bob's baked goods","bob's bakery","bakery website","bakery web site","bakery home page"),
                ("Mobile App for Pizza Place","pizza app","pizza mobile app","jane janeson","mobile order")
```
In this example, 2 projects will be analyzed (relative to each other). Keywords associated with each are defined accordingly. The script is extensible so key words can include descriptive text, names of those associated with a project, etc.

### Using properties
May apply if you manage or own properties (real estate)
```
$keywordList =  ("Sycamore Property","2897 sycamore st","sycamore plaza","98123135135"),
                ("Elm Property","elm","elm center","elm apartment"),
                ("Downtown Property","1 civic plaza","downtown office","9889231"),
                ("Santa Fe Property","santa fe","plaza")
 ```
 In this case, keywords are property names, addresses, and/or id numbers.
