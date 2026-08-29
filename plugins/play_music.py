"""
Spotify Play Music Plugin
Searches and plays music on Spotify using URI handlers and UI automation.
"""

import os
import time
import urllib.parse
try:
    import pyautogui
except ImportError:
    pyautogui = None

PLUGIN = {
    "name": "control_music",
    "description": (
        "Use this tool to play a specific song on Spotify, or to pause/stop current music. "
        "Do NOT use open_app for playing songs; use this tool instead."
    ),
    "parameters": {
        "type": "OBJECT",
        "properties": {
            "action": {
                "type": "STRING",
                "description": "The action to perform: 'play', 'pause', or 'stop'."
            },
            "song_name": {
                "type": "STRING", 
                "description": "The name of the song, artist, or playlist to search and play. Leave empty if just pausing/stopping."
            },
        },
        "required": ["action"],
    },
}

def run(parameters: dict, player=None, session_memory=None) -> str:
    action = parameters.get("action", "play").lower()
    song_name = parameters.get("song_name", "").strip()
    
    if not pyautogui:
        return "pyautogui is not installed. Please pip install pyautogui to use Spotify automation."

    try:
        if action == "pause" or action == "stop":
            if player:
                player.write_log(f"SAM: {action.capitalize()}ing music...")
            
            # Press media play/pause key which stops/pauses media in Windows
            pyautogui.press('playpause')
            return f"I have {action}ed the music."

        # If action is 'play'
        if song_name:
            if player:
                player.write_log(f"SAM: Searching Spotify for '{song_name}'...")
            
            # Format the Spotify search URI
            query = urllib.parse.quote(song_name)
            uri = f"spotify:search:{query}"
            
            # Open Spotify with the search query
            os.startfile(uri)
            
            # Wait for Spotify to open and load the search results
            time.sleep(3.0)
            
            # Navigate to play button
            pyautogui.press('tab', presses=3, interval=0.1)
            pyautogui.press('enter')
            
            return f"I've opened Spotify and started playing {song_name}."
            
        else:
            if player:
                player.write_log("SAM: Resuming music on Spotify...")
            
            # Open Spotify to bring it to front
            os.startfile("spotify:")
            time.sleep(1.0)
            
            # Press media play/pause key
            pyautogui.press('playpause')
            return "I've resumed your music."
            
    except Exception as e:
        return f"Sorry, I encountered an error trying to {action} music: {str(e)}"
