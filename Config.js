/**
 * Configuration and script properties for the Email Assistant add-on.
 * All URLs and secrets are loaded from Script Properties (Project Settings).
 */
var SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
var SERVER_DOMAIN = SCRIPT_PROPERTIES.getProperty('SERVER_DOMAIN');
var FRONTEND_URL = SCRIPT_PROPERTIES.getProperty('FRONTEND_URL') || 'https://replai.us';
