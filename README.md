# Outlook-Link-to-Trello

A VBA module to create a bi-directional link between Outlook email items and Trello.

## Project Status
**This project is incomplete at this time.**
I have created a Product Management Dashboard on Notion.so to keep track of the current status and feature roadmap.
You can find this resource at [Outlook Link to Trello | Product Dashboard](https://jumpy-catfish-a5f.notion.site/Outlook-Link-to-Trello-a08953871df6443ba7e3b8eb7b1b923a). 

## Problem Statement
Official versions of "Outlook to Trello" integrations lack a very important feature: the ability to open the trigger email message in Outlook *from Trello*. This module adds that functionality, creating a static link between the email and the Trello card that enables the user to open Outlook mail items directly from Trello.

## History
My initial implementation of this system used VBA to extract email information and concatenate it into a string that would be saved to the clipboard. This VBA module was triggered by an AutoHotKey script that would use alt-hotkeys to run the Macro from the Outlook GUI, wait for the clipboard to be filled, then would trigger a Python script that would parse the data on the clipboard, and handle the HTTP requests to interface with Trello.
As one may imagine, this mixed software stack was the source of many headaches.

## INI Schema
This project uses an ini file to enable data persistance. The schema can be found below:

### Sections
- [app]
 - first-run-complete = Boolean (safe to convert with CBool function)
- [trello]
 - api_key = API key
 - api_token = API token
 - user_id = Trello user ID for API requests
 - list_id = ListID of the target list for new cards