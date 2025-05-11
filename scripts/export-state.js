#!/usr/bin/env node

/**
 * Legacy Sheet State Export Tool
 * 
 * This script exports the current application state to a shareable file.
 * It can also import a state file from another user.
 * 
 * Usage:
 *   node export-state.js --export [optional_filename]  # Export current state
 *   node export-state.js --import <filename>          # Import state from file
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

// Get the directory of the current script
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Get the root directory of the project
const rootDir = path.resolve(__dirname, '..');
const dataDir = path.join(rootDir, 'data');
const appStateFile = path.join(dataDir, 'appState.json');

// Make sure the data directory exists
if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
  console.log(`Created data directory: ${dataDir}`);
}

// Parse command line arguments
const args = process.argv.slice(2);

if (args.includes('--export')) {
  // Export the current state
  exportState(args[args.indexOf('--export') + 1]);
} else if (args.includes('--import')) {
  // Import a state file
  const importFile = args[args.indexOf('--import') + 1];
  if (!importFile) {
    console.error('Error: No import file specified');
    printUsage();
    process.exit(1);
  }
  importState(importFile);
} else {
  // Print usage instructions
  printUsage();
}

/**
 * Export the current application state to a file
 * @param {string} [customFilename] - Optional custom filename
 */
function exportState(customFilename) {
  try {
    // Check if the state file exists
    if (!fs.existsSync(appStateFile)) {
      console.error('Error: No application state found. Run the application and save state first.');
      process.exit(1);
    }

    // Read the current state
    const state = fs.readFileSync(appStateFile, 'utf8');
    
    // Parse it to get metadata and validate
    const stateObj = JSON.parse(state);
    
    // Generate a default filename if none provided
    let exportFilename = customFilename;
    if (!exportFilename) {
      const date = new Date().toISOString().replace(/:/g, '-').replace(/\..+/, '');
      exportFilename = `legacy-sheet-state-${date}.json`;
    }
    
    // Make sure the filename ends with .json
    if (!exportFilename.endsWith('.json')) {
      exportFilename += '.json';
    }
    
    // Create the export path
    const exportPath = path.join(process.cwd(), exportFilename);
    
    // Write the state to the export file
    fs.writeFileSync(exportPath, state);
    
    // Add metadata to console output
    const pagesCount = stateObj.pages ? stateObj.pages.length : 0;
    const savedDate = stateObj.savedAt ? new Date(stateObj.savedAt).toLocaleString() : 'unknown';
    
    console.log(`\nState exported successfully!`);
    console.log(`- File: ${exportPath}`);
    console.log(`- Pages: ${pagesCount}`);
    console.log(`- Saved on: ${savedDate}`);
    console.log(`\nShare this file with others to continue editing from where you left off.`);
    
    process.exit(0);
  } catch (error) {
    console.error(`Error exporting state: ${error.message}`);
    process.exit(1);
  }
}

/**
 * Import a state file into the application
 * @param {string} importFile - Path to the state file to import
 */
function importState(importFile) {
  try {
    // Resolve the import file path
    const importPath = path.resolve(process.cwd(), importFile);
    
    // Check if the import file exists
    if (!fs.existsSync(importPath)) {
      console.error(`Error: Import file not found: ${importPath}`);
      process.exit(1);
    }
    
    // Read the import file
    const state = fs.readFileSync(importPath, 'utf8');
    
    // Parse it to validate and get metadata
    const stateObj = JSON.parse(state);
    
    // Create a backup of the current state if it exists
    if (fs.existsSync(appStateFile)) {
      const backupFile = `${appStateFile}.backup-${Date.now()}`;
      fs.copyFileSync(appStateFile, backupFile);
      console.log(`Created backup of current state: ${backupFile}`);
    }
    
    // Write the import file to the application state file
    fs.writeFileSync(appStateFile, state);
    
    // Add metadata to console output
    const pagesCount = stateObj.pages ? stateObj.pages.length : 0;
    const savedDate = stateObj.savedAt ? new Date(stateObj.savedAt).toLocaleString() : 'unknown';
    
    console.log(`\nState imported successfully!`);
    console.log(`- Pages: ${pagesCount}`);
    console.log(`- Originally saved on: ${savedDate}`);
    console.log(`\nRun the application to continue editing from this state.`);
    
    process.exit(0);
  } catch (error) {
    console.error(`Error importing state: ${error.message}`);
    process.exit(1);
  }
}

/**
 * Print usage instructions
 */
function printUsage() {
  console.log(`
Legacy Sheet State Export/Import Tool
------------------------------------

This tool helps you export and import application state to share with others.

Usage:
  node export-state.js --export [optional_filename]  # Export current state
  node export-state.js --import <filename>          # Import state from file

Examples:
  node export-state.js --export                     # Export with auto-generated filename
  node export-state.js --export my-project.json     # Export with specific filename
  node export-state.js --import my-project.json     # Import from specific file
`);
} 
