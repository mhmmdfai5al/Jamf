#! /bin/bash

# Extensions Atrribute: Excel was installed from the AppStore

# Check if Application is from the apps store
# written by Faisal Mohammad 2nd DEC 2022

# Q. Is the app installed from the App Store?
# Return values are YES | NO
# YES means all the Application is installed from Apple App Store. NO is self explanatory.

# Checking Microsoft Excel

APPSTORE=$(/usr/bin/mdfind "kMDItemAppStoreHasReceipt=1" | grep "Microsoft Excel")

if [[ ${APPSTORE} ]]; then
	echo "YES"
else 
	ALLAPPSTORE='NO'
	echo "${ALLAPPSTORE}"
fi