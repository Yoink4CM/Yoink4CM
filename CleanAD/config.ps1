$DaysInactive = 730  #How many days should no one have logged into a PC before it's disabled?
$SearchOU = "OU=SearchOU,OU=Computer-Accounts,DC=yourdomain,DC=com" #which OU would you like to search for unused computers?
$FilteredOU = "OU=FilteredOU,OU=Computer-Accounts,DC=yourdomain,DC=com" #which OU would you like to move unused computers to?