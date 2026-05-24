/**
 * Error Logger - Capture and log runtime errors
 * This script should be included in the app to capture JavaScript errors
 */

// Create a simple error log that can be viewed in the browser console
window.APP_ERROR_LOG = [];
window.APP_WARNINGS = [];

// Capture uncaught errors
window.addEventListener('error', (event) => {
  const errorInfo = {
    timestamp: new Date().toISOString(),
    message: event.message,
    filename: event.filename,
    line: event.lineno,
    column: event.colno,
    stack: event.error?.stack || 'No stack trace',
    type: 'UNCAUGHT_ERROR',
  };

  window.APP_ERROR_LOG.push(errorInfo);

  console.error('🔴 UNCAUGHT ERROR:', {
    message: event.message,
    file: `${event.filename}:${event.lineno}:${event.colno}`,
    stack: event.error?.stack,
  });
});

// Capture unhandled promise rejections
window.addEventListener('unhandledrejection', (event) => {
  const errorInfo = {
    timestamp: new Date().toISOString(),
    message: event.reason?.message || String(event.reason),
    stack: event.reason?.stack || 'No stack trace',
    type: 'UNHANDLED_PROMISE_REJECTION',
  };

  window.APP_ERROR_LOG.push(errorInfo);

  console.error('🔴 UNHANDLED PROMISE REJECTION:', event.reason);
});

// Override console.error to track errors
const originalError = console.error;
console.error = function(...args) {
  const errorInfo = {
    timestamp: new Date().toISOString(),
    message: args.map(arg => typeof arg === 'string' ? arg : JSON.stringify(arg)).join(' '),
    type: 'CONSOLE_ERROR',
  };

  window.APP_ERROR_LOG.push(errorInfo);
  originalError.apply(console, args);
};

// Override console.warn to track warnings
const originalWarn = console.warn;
console.warn = function(...args) {
  const warning = {
    timestamp: new Date().toISOString(),
    message: args.map(arg => typeof arg === 'string' ? arg : JSON.stringify(arg)).join(' '),
  };

  window.APP_WARNINGS.push(warning);
  originalWarn.apply(console, args);
};

// Function to display error log in browser
window.showErrorLog = function() {
  console.group('📋 ERROR LOG');
  if (window.APP_ERROR_LOG.length === 0) {
    console.log('✅ No errors logged');
  } else {
    window.APP_ERROR_LOG.forEach((error, index) => {
      console.group(`Error #${index + 1} - ${error.type}`);
      console.log('Time:', error.timestamp);
      console.log('Message:', error.message);
      if (error.filename) console.log('File:', error.filename);
      if (error.line) console.log('Line:Column', `${error.line}:${error.column}`);
      if (error.stack) console.log('Stack:', error.stack);
      console.groupEnd();
    });
  }
  console.groupEnd();
};

// Function to display warnings in browser
window.showWarnings = function() {
  console.group('⚠️  WARNING LOG');
  if (window.APP_WARNINGS.length === 0) {
    console.log('✅ No warnings logged');
  } else {
    window.APP_WARNINGS.forEach((warning, index) => {
      console.log(`Warning #${index + 1}:`, warning.message);
    });
  }
  console.groupEnd();
};

// Function to export logs
window.exportLogs = function() {
  const logs = {
    errors: window.APP_ERROR_LOG,
    warnings: window.APP_WARNINGS,
    exportedAt: new Date().toISOString(),
  };

  console.log('Downloading logs...');
  const blob = new Blob([JSON.stringify(logs, null, 2)], { type: 'application/json' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `error-log-${Date.now()}.json`;
  document.body.appendChild(a);
  a.click();
  window.URL.revokeObjectURL(url);
  document.body.removeChild(a);
};

// Display help
window.showDebugHelp = function() {
  console.log(`
╔════════════════════════════════════════════════════════════╗
║              DEBUG HELPER COMMANDS                         ║
╚════════════════════════════════════════════════════════════╝

Available commands in browser console:

showErrorLog()     - Display all logged errors
showWarnings()     - Display all warnings
exportLogs()       - Download error log as JSON

Properties:
APP_ERROR_LOG      - Array of all errors
APP_WARNINGS       - Array of all warnings

Quick reference:
- Open DevTools: F12 or Ctrl+Shift+I
- Go to Console tab
- Type command and press Enter
  `);
};

// Show info on app start
console.log('%c✓ Error Logger Initialized', 'color: green; font-weight: bold;');
console.log('Type %cshowDebugHelp()%c to see available debug commands', 'color: blue;', 'color: auto;');

// Press F12 → Go to "Console" tab
