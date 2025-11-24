# Component Label Maker

## Overview
The purpose of this program is to create a scannable label for parts whose original DataMatrix code is unreadable, non-existent, or missing. The original [DigiKey Organizer](https://github.com/grossrc/DigiKey_Organizer) program was created to provide a smooth/convenient method for storing, browsing, and using parts. This manual labelling scheme extends that same convenience to parts without an explicit scannable code. It also reduces the likelihood of improper entry into the storage system by having the code, quantity, and part number printed directly onto the storage bag.

![alt text](assets/Program%20Still.png)

## Demo Usage
[![Watch the video](https://img.youtube.com/vi/nTQUwghvy5Q/default.jpg)](https://youtu.be/nTQUwghvy5Q)

## Requirements
Built on Python 3.13.2. Interfaces with a NiimBot B1 label maker. The program is built to accommodate other models of Niimbot printers, but they have not been explicitly tested.

## Quick Start

1.  **Get the Code**
    *   Click the green **Code** button at the top of this page and select **Download ZIP**.
    *   Extract the ZIP file to a folder on your computer.
    *   Open a terminal (Command Prompt or PowerShell) by right clicking the folder. This will open with the proper directory.

2.  **Install Dependencies**
    Run the following command:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Configure DigiKey API**
    *   Register an app at the [DigiKey API Portal](https://developer.digikey.com/) to get your Client ID and Secret.
    *   Copy the example configuration file to a new file named `.env`:
        ```bash
        copy .env.example .env
        ```
    *   Open the new [`.env`](.env ) file and paste your credentials.

4.  **Run the App**
    ```bash
    python label_maker_app.py
    ```

## DigiKey API Help
Getting the DigiKey API credentials can be difficult. I clicked around for a while before finding out where I needed to go. A Client_ID and Client_Secret are what's used by the DigiKey API to verify queries. To get these parameters from your DigiKey account, you must go to https://developer.digikey.com/ and create an organization/project/production app. The only specific API endpoint used for this program is "ProductInformation V4", so make sure that is what's selected when creating your "production app". Once created, you should have credentials which can be copied over to your .env file. This file will be referenced by the program anytime your credentails are needed.