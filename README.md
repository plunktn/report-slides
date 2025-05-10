# Report Slides Generator

This repository contains a script to generate Google Slides reports for different brands based on a JSON configuration file. The script reads the configuration from `assets.json` and creates a Google Slides presentation for each brand, customizing it with logos, ads, and other details.

## Features
- Reads configuration from `assets.json`.
- Generates a Google Slides presentation for each brand.
- Customizes slides with brand logos, ads, and background images.

## Prerequisites
1. **Node.js**: Ensure you have Node.js installed on your system.
2. **Google API Credentials**: A service account JSON key file is required to authenticate with Google Slides and Drive APIs. Place the key file in the root directory and update the script to use it.
3. **Google Slides Template**: Ensure the `template_id` in `assets.json` points to a valid Google Slides template.

## Setup
1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd report-slides
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Place your Google API credentials file (e.g., `service-account.json`) in the root directory.
4. Update the `auth` section in `template.js` to use your credentials file.

## Running the Script
1. Ensure `assets.json` is correctly configured with the required data.
2. Run the script:
   ```bash
   node template.js
   ```
3. The script will generate a Google Slides presentation for each brand listed in `assets.json`.

## File Structure
- `assets.json`: Configuration file containing details about the client, brands, and ads.
- `template.js`: Main script to generate Google Slides presentations.
- `script.js`: (Optional) Additional scripts or utilities.
- `clientes/`: Directory for client-specific data.

## Example
An example `assets.json` is provided in the repository. Update it with your data to generate reports.

## Notes
- Ensure the `template_id` in `assets.json` is accessible by the service account.
- The generated presentations will be saved in your Google Drive.