#-------------------------------------------------------------------------------
# Name:        Wirecast COM interface Object (Windows)
# Purpose:     Communicate with Wirecast API 
# Author:      Benjamin Yaroch
# Created:     03/23/2015
# Copyright:   (c) DynamoBen 2015
#
# Uses: http://sourceforge.net/projects/pywin32/
#
# An easier way to use this module is to install ActiveState ActivePython...
# ...doing so includes all the COM dependancies.
# http://www.activestate.com/activepython
#
# This module is also threading friendly.
#
# 3/23/2015 - v1.0 Initial Release.
#
#-------------------------------------------------------------------------------

import win32com.client, pythoncom

#=======================================================
# Wirecast Interface
#=======================================================
def DocumentByName(name, compareMethod):
    try:
        pythoncom.CoInitialize() 
        objWirecast = win32com.client.GetActiveObject("Wirecast.Application")
        if objWirecast:
            objDoc = objWirecast.DocumentByIndex(name, compareMethod)
            return objDoc
    except:
        #pass
        print("Wirecast app not found.")
        return 0

def DocumentByIndex(idx):
    try:
        pythoncom.CoInitialize() 
        objWirecast = win32com.client.GetActiveObject("Wirecast.Application")
        if objWirecast:
            objDoc = objWirecast.DocumentByIndex(idx)
            return objDoc
    except:
        #pass
        print("Wirecast app not found.")
        return 0

#=======================================================
# Document Interface
#=======================================================
        
    # ==== Methods ====
def Broadcast(status):
    # "start" - start broadcasting. If already broadcasting, nothing happens.
    # "stop" - stop broadcasting. If not broadcasting, nothing happens.
    # Does not return a result. 
    objDoc = DocumentByIndex(1)
    if objDoc:    
        objDoc.Broadcast(status)                

def IsBroadcasting():
    # Returns 1 if Wirecast is currently broadcasting.
    # Returns 0 if Wirecast is not currently broadcasting.
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.IsBroadcasting

def ArchiveToDisk(status):
    # "start" - start archiving to disk. If already archiving to disk, nothing happens.
    # "stop" - stop archiving to disk. If not archiving to disk, nothing happens.
    # Does not return a result. 
    objDoc = DocumentByIndex(1)
    if objDoc:
        objDoc.ArchiveToDisk(status)            

def IsArchiveToDisk():
    # Returns 1 if Wirecast is currently archiving to disk.
    # Returns 0 if Wirecast is not currently archiving to disk. 
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.IsArchivingToDisk()

def LayerByIndex(layerIdx):
    # The 1-based index (1..N) of the document layer to retrieve.
    # Returns a Wirecast Document Layer Object.
    # Return 0 if index is out of bounds, or document cannot be found. 
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.LayerByIndex(layerIdx)    

def LayerByName(layerName):
    # The current name of the layer in Wirecast.
    # Returns a Wirecast Document Layer Object.
    # Return 0 if index is out of bounds, or document cannot be found. 
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.LayerByIndex(layerName)     
    
def ShotByShotID(shotID):
    # An ID retrieved from various functions, such as ShotIDByName and ShotIDByName.
    # Returns a Wirecast Shot Object.
    # Return 0 if index is out of bounds, or document cannot be found. 
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.ShotByShotID(shotID)        

def ShotIDByShotName(shotName):
    # The string to use to compare against each shot name.
    # The comparison method, which is one of these values:
    #   0 = Exact Match 
    #   1 = Contains 
    #   2 = Case Insensitive
    #   3 = Case Insensitive Contains
    # Returns a Wirecast shot_id.
    # Return 0 if no shot matches.    
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.ShotIDByName(shotName)      
    
def SaveSnapshot(pathString):
    # Path (including file name) to file where snapshot should be saved.
    objDoc = DocumentByIndex(1)
    if objDoc:
        objDoc.SaveSnapshot(pathString)

def RemoveMedia(pathString):
    # Path (including file name) to file that should be removed.
    # Does not return a result. 
    objDoc = DocumentByIndex(1)
    if objDoc:
        objDoc.RemoveMedia(pathString)    

    # ==== Properties ====
def TransitionSpeed(speed):
    # TransitionSpeed is a string.
    # This property maps to the menu item "Switch --> Transition Speed". 
    # Speeds: "slowest", "slow", "normal", "fast", "fastest"
    objDoc = DocumentByIndex(1)
    if objDoc:
        objDoc.TransitionSpeed = speed

def getTransitionSpeed():		        
    # Gets current TransitionSpeed, returns as a string.
    # Speeds: "slowest", "slow", "normal", "fast", "fastest"    
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.TransitionSpeed
    
def AutoLive(state):                        
    # AutoLive is an int.
    # This property maps to the AutoLive checkbox which can be found...
    # ...at the bottom of the document window in Wirecast. 
    # 1=On, 0=Off
    objDoc = DocumentByIndex(1)
    if objDoc:
        objDoc.AutoLive = state                 
        
def getAutoLive():
    # Gets current AutoLive, returns as an int.
    # 1=On, 0=Off    
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.AutoLive

def ActiveTransitionIndex(idx):
    # ActiveTransitionIndex is an int.
    # This property defines which of the two transition poups buttons are active.
    # 1 = The first popup is active, 2 = The second popup is active
    objDoc = DocumentByIndex(1)
    if objDoc:
        objDoc.ActiveTransitionIndex = idx

def getActiveTransitionIndex():
    # Get current ActiveTransitionIndex,return as an int.
    # 1 = The first popup is active, 2 = The second popup is active
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.ActiveTransitionIndex          

def AudioMutedToSpeaker(state):
    # AudioMutedToSpeaker is an int.
    # This property defines whether the output the headphones is muted or not.
    # 1=Muted, 0=Unmuted/Audible
    objDoc = DocumentByIndex(1)
    if objDoc:
        objDoc.AudioMutedToSpeaker = state      

def getAudioMutedToSpeaker():
    # Get current AudioMutedToSpeaker, return as an int.
    # 1=Muted, 0=Unmuted/Audible    
    objDoc = DocumentByIndex(1)
    if objDoc:
        return objDoc.AudioMutedToSpeaker     

#=======================================================
# Layer Interface
#=======================================================

    # ==== Methods ====
def ShotCount(layerNum):
    # Returns the number of shots that are currently on the document layer. 
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        return objLayer.ShotCount()
    
def ShotIDByIdx(layerNum, shotIdx):
    # The 1-based index (1..N) of the document to retrieve.
    # Returns a shot_id, or zero on error (index out of range, etc). 
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        return objLayer.ShotIDByIndex(shotIdx)    
    
def ShotIDByName(layerNum, shotName, compareMethod):
    # The string to use to compare against each shot name.
    # The comparison method, which is one of these values:
    #   0 = Exact Match 
    #   1 = Contains 
    #   2 = Case Insensitive
    #   3 = Case Insensitive Contains
    # Returns a Wirecast shot_id.
    # Return 0 if no shot matches.
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        return objLayer.ShotIDByName(shotName, compareMethod)

def PreviewShotID(layerNum):
    # Returns a Wirecast shot_id, which can be used in various methods such as ShotByShotID
    # Return 0 if no shot matches. 
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        return objLayer.PreviewShotID()             # ShotID of shot in preview

def LiveShotID(layerNum):
    # Returns a Wirecast shot_id, which can be used in various methods such as ShotByShotID
    # Return 0 if no shot matches
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        return objLayer.LiveShotID()                # ShotID of shot in live

def AddShotWithMedia(layerNum, pathString):
    # The full path to the media file on disk.
    # Returns a Wirecast shot_id, which can be used in various methods such as ShotByShotID
    # Return 0 if media cannot be loaded.
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        return objLayer.AddShotWithMedia(pathString)
        
def RemoveShotByID(layerNum, shotID):
    # An ID retrieved from various functions, such as ShotIDByName and ShotIDByName. 
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        objLayer.RemoveShotByID = shotID   
    
def Go(layerNum):
    # This brings the current state of the document layer live. 
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        objLayer.Go()

    # ==== Properties ====
def Visible(layerNum, state):
    # Visible is an int. 
    # This property defines the visibility of a layer.
    # 1 = Visible, 0 = Not visible
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        objLayer.Visible = state

def getVisible(layerNum):
    # Gets current Visible state, returned as an int. 
    # 1 = Visible, 0 = Not visible
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        return objLayer.Visible      
    
def ActiveShotID(layerNum, shotID):
    # ActiveShotID is an int. 
    # Any shot_id returned from the document layer functions.
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        objLayer.ActiveShotID = shotID

def getActiveShotID(layerNum):
    # Gets current ActiveShotID, returns as an int. 
    objLayer = LayerByIndex(layerNum)
    if objLayer:
        return objLayer.ActiveShotID      
    
#=======================================================
# Shot Interface
#=======================================================
        
    # ==== Methods ====
def Preview(shotID):
    # Whether the shot is being displayed in the Preview view
    # 0 = Not in Preview, 1 = In Preview
    objShot = ShotByShotID(shotID)
    if objShot:
        return objShot.Preview()                        

def Live(shotID):
    # Whether the shot is being displayed Live
    # 0 = Not Live , 1 = Live   
    objShot = ShotByShotID(shotID)
    if objShot:
        return objShot.Live()

def Playlist(shotID):
    # Returns whether or not the shot is a Playlist.
    # 0 = Not in Playlist , 1 = In Playlist   
    objShot = ShotByShotID(shotID)
    if objShot:
        return objShot.Playlist()

def NextShot(shotID):
    # If the shot object is a Playlist and live, transitions to the Next shots. 
    objShot = ShotByShotID(shotID)
    if objShot:
        objShot.NextShot()

def PreviousShot(shotID):
    # If the shot object is a Playlist and live, transitions to the Previous shots.  
    objShot = ShotByShotID(shotID)
    if objShot:
        objShot.PreviousShot()         

    # ==== Properties ====
def Name(shotID, name):
    # Name is a String.
    # Setting the name is functionally equivalent to selecting "Rename Shot."
    objShot = ShotByShotID(shotID)
    if objShot:
        objShot.Name = name

def getName(shotID):
    # Name is a String.
    # Setting the name is functionally equivalent to selecting "Rename Shot."
    objShot = ShotByShotID(shotID)
    if objShot:
        return objShot.Name
    
