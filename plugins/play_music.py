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
    "name": "play_music",
    "description": (
        "Use this tool to play a specific song on Spotify, or just play/resume music. "
        "Do NOT use open_app for playing songs; use this tool instead."
    ),
    "parameters": {
        "type": "OBJECT",
        "properties": {
            "song_name": {
                "type": "STRING", 
                "description": "The name of the song, artist, or playlist to search and play. Leave empty to just resume current music."
            },
        },
        "required": [],
    },
}

def run(parameters: dict, player=None, session_memory=None) -> str:
    song_name = parameters.get("song_name", "").strip()
    
    if not pyautogui:
        return "pyautogui is not installed. Please pip install pyautogui to use Spotify automation."

    try:
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
            
            # In the Spotify desktop app, after a search URI is opened, 
            # pressing Tab a few times and Enter will usually play the top result.
            # We hit Tab 3 times to get to the Top Result play button and hit Enter.
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.press('enter')
            
            return f"I've opened Spotify and started playing {song_name}."
            
        else:
            if player:
                player.write_log("SAM: Resuming music on Spotify...")
            
            # Just open Spotify to bring it to front
            os.startfile("spotify:")
            time.sleep(1.0)
            
            # Press media play/pause key
            pyautogui.press('playpause')
            
            return "I've resumed your music on Spotify."
            
    except Exception as e:
        return f"Sorry, I encountered an error trying to play music: {str(e)}"
