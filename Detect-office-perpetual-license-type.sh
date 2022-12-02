#!/bin/zsh
#set -x

# Extensions Attribute: Office License Type e.g. Office 365 / Perpetual (2016, 2019, 2021, etc)

# Check if Office Application license type is perpetual OR subscription based
# written by Faisal Mohammad 22nd NOV 2022

# Q. What type of Office license installed?
# Return values are License description encoded in "/Library/Preferences/com.microsoft.office.licensingV2.plist" 

# Constants
SetConstants() {
	O365PRODUCT="$HOME/Library/Group Containers/UBF8T346G9.Office"
	WORDAPPPATH="/Applications/Microsoft Word.app"
	PERPETUALLICENSE="/Library/Preferences/com.microsoft.office.licensingV2.plist"
	SHAREDLICENSE="/Library/Application Support/Microsoft/Office365/com.microsoft.Office365.plist"
	REGISTRY="$O365PRODUCT/MicrosoftRegistrationDB.reg"
	VNEXTPERPETUALLICENSEPATH="/Library/Microsoft/Office/Licenses"
}

# Checks to see if a volume license file is present
DetectPerpetualLicense() {
	if [ -f "$PERPETUALLICENSE" ] || [ -e "$VNEXTPERPETUALLICENSEPATH" ]; then
		echo "1"
	else
		echo "No perpetual license found"
	fi
}

# Reports the type of perpetual license installed
PerpetualLicenseType() {
	if [ -f "$PERPETUALLICENSE" ]; then
		/bin/cat "$PERPETUALLICENSE" | grep -q "A7vRjN2l/dCJHZOm8LKan11/zCYPCRpyChB6lOrgfi"
		if [ "$?" = "0" ]; then
				echo "Office 2019 Volume License"
				return
		fi
		/bin/cat "$PERPETUALLICENSE" | grep -q "Bozo+MzVxzFzbIo+hhzTl4JKv18WeUuUhLXtH0z36s"
		if [ "$?" = "0" ]; then
				echo "Office 2019 Preview Volume License"
				return
		fi
		/bin/cat "$PERPETUALLICENSE" | grep -q "A7vRjN2l/dCJHZOm8LKan1Jax2s2f21lEF8Pe11Y+V"
		if [ "$?" = "0" ]; then
				echo "Office 2016 Volume License"
				return
		fi
		/bin/cat "$PERPETUALLICENSE" | grep -q "DrL/l9tx4T9MsjKloHI5eX"
		if [ "$?" = "0" ]; then
				echo "Office 2016 Home and Business License"
				return
		fi
		/bin/cat "$PERPETUALLICENSE" | grep -q "C8l2E2OeU13/p1FPI6EJAn"
		if [ "$?" = "0" ]; then
				echo "Office 2016 Home and Student License"
				return
		fi
		/bin/cat "$PERPETUALLICENSE" | grep -q "Bozo+MzVxzFzbIo+hhzTl4m"
		if [ "$?" = "0" ]; then
				echo "Office 2019 Home and Business License"
				return
		fi
		/bin/cat "$PERPETUALLICENSE" | grep -q "Bozo+MzVxzFzbIo+hhzTl4j"
		if [ "$?" = "0" ]; then
				echo "Office 2019 Home and Student License"
				return
		fi
		echo "Office Perpetual License"
	fi
	if [ -e "$VNEXTPERPETUALLICENSEPATH" ]; then
		echo "Office 2021 Consumer License"
		return
	fi
}