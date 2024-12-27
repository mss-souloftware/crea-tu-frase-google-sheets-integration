<?php
/**
 * Plugin Name: Google Sheets Integration (Crea Tu Frase Add-On) 
 * Plugin URI: https://souloftware.com/
 * Description: Connect Google Sheets with Crea Tu Frase for order management.
 * Version: 1.0.0
 * Author: Souloftware
 * Author URI: https://souloftware.com/contact
 */

if (!defined('ABSPATH')) {
    exit; // Exit if accessed directly
}

// Include required files
require_once plugin_dir_path(__FILE__) . 'includes/class-google-auth.php';
require_once plugin_dir_path(__FILE__) . 'includes/class-google-sheets.php';

// Initialize the plugin
function gsi_init()
{
    new Google_Auth();
}
add_action('plugins_loaded', 'gsi_init');
