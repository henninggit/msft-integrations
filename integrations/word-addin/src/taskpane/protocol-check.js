// Add this to your taskpane.js to detect and warn about HTTP in production
if (location.protocol === 'http:' && location.hostname !== 'localhost') {
    console.warn('⚠️  Office add-in running on HTTP in production - this will not work!');
    console.warn('Production Office add-ins MUST use HTTPS with valid certificates');
}

// For development, show protocol info
console.log(`🔧 Add-in loaded via ${location.protocol}//${location.host}`);
console.log('🔧 Development mode:', location.hostname === 'localhost');
