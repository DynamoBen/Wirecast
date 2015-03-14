#-------------------------------------------------------------------------------
# Name:        WCsnap2social
# Purpose:     Post Wirecast snapshots to social media
#
# Author:      Benjamin Yaroch
#
# Created:     06/11/2014 - Initial Release
#
#-------------------------------------------------------------------------------

import os, glob, re, win32com.client 
from twython import Twython

snapshotPath    = 'PUT_THE_PATH_TO_THE_IMAGES_FOLDER_HERE'
message         = 'PUT MESSAGE YOU WANT TO TWEET HERE'

# === Twitter credentials ===
CONSUMER_KEY    = 'PUT_YOUR_CONSUMER_SECRET_HERE'
CONSUMER_SECRET = 'PUT_YOUR_CONSUMER_SECRET_HERE'
OAUTH_TOKEN     = 'PUT_YOUR_OAUTH_TOKEN_HERE'
OAUTH_SECRET    = 'PUT_YOUR_OAUTH_SECRET_HERE'

#=======================================================
# Functions Calls
#=======================================================
def post2twitter(pathString, msg):		        # Twitter Post with Image

    photo = open(pathString, 'rb')                      # Open the file to be posted

    # Send to twitter with message and image
    twitter = Twython(CONSUMER_KEY,CONSUMER_SECRET,OAUTH_TOKEN,OAUTH_SECRET) # Authenticate
    twitter.update_status_with_media(media=photo, status=msg)               # Post

def SaveSnapshot(pathString):                           # Gets snapshot image from Wirecast and saves it
    currentImages = glob.glob(pathString + "*.png")     # Look for any existing .png images

    numList = [0]
    for img in currentImages:
        i = os.path.splitext(img)[0]                    # Split the path prior to searching
        try:
            num = re.findall('[0-9]+$', i)[0]           # Search for all images with the a similar name
            numList.append(int(num))                    # Store those file names in a list
        except IndexError:
            pass
    numList = sorted(numList)                           # Sort the list numerically
    newNum = numList[-1]                                # Get the next number in the series    
    saveName = 'snapshot_%04d.png' % newNum             # Determine next filename (IE snapshot_0001.png)

    objWirecast = win32com.client.Dispatch("Wirecast.Application") # Connect to Wirecast 
    if objWirecast:                                     # If successful...
        objDoc = objWirecast.DocumentByIndex(1)
        objDoc.SaveSnapshot(pathString + saveName)      # ...take a snapshot and store it with path and filename
        return pathString+saveName                      # Return path and new filename
    
#=======================================================
# Main
#=======================================================
path = SaveSnapshot(snapshotPath)               # Gets snapshot image
print(path)                    
post2twitter(path, message)                     # Sends to twitter
print "Success! Posted to Twitter"


