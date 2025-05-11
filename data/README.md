# Legacy Sheet App State Storage

This directory contains the persistent state files for the Legacy Sheet Genealogy Document Generator application.

## Files

- `appState.json`: Contains the saved state of the application, including all formatted pages, generation titles, and other content.

## Purpose

This directory enables state persistence across different instances of the application. When multiple users work on the same codebase, the state will be preserved and loaded, allowing them to continue from where the previous user left off.

## Usage

The state is automatically saved when:
1. You click the "Save State" button in the pagination controls
2. You close the application
3. Every 5 minutes while the application is running

Do not manually edit these files unless you understand the state structure.

## Troubleshooting

If you encounter issues with the state loading or saving:
1. Delete the `appState.json` file to reset to the default template
2. Make sure the `data` directory has write permissions 
