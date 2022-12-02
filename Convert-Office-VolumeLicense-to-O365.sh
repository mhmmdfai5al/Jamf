#!/bin/bash

# Converts office perpetual license to office 365 subscription type (user involvement required) 

# Kills all running office applications and then deleted license info PLIST file "/Library/Preferences/com.microsoft.office.licensingV2.plist" 
# written by Faisal Mohammad 12th OCT 2022

#NOTE: This script only remove the license file, conversion happens once users sign in using M365 credentials

license=/Library/Preferences/com.microsoft.office.licensingV2.plist
#echo $license

if [ -f $license ]
then
	# Change permission of license file
	chmod 750 $license

	# Relative process kill by Keyword | use "man pkil" to know more
	# Relevant processes can be dimmed/enabled as per requirement related to your use case/scenario 
	echo "Killing relevant Microsoft Office Process..."
	pkill -f "Microsoft Excel"
	pkill -f "Microsoft Word"
	pkill -f "Microsoft Outlook"
	pkill -f "Microsoft Powerpoint"
	pkill -f "Microsoft Teams"
	pkill -f "OneDrive"
	pkill -f "OneNote"
	#Remove Office perpetual license file plist
	echo "Deleting $license file"
	rm $license
	open /Applications/Microsoft\ Word.app
fi