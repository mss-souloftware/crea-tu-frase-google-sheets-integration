<?php

class Google_Auth
{
    private $client;

    public function __construct()
    {
        add_action('admin_menu', [$this, 'add_admin_page']);
        add_action('admin_init', [$this, 'handle_oauth']);
        $this->setup_client();
    }

    private function setup_client()
    {
        require_once __DIR__ . '/lib/vendor/autoload.php';

        $this->client = new Google_Client();
        $this->client->setClientId('YOUR_CLIENT_ID');
        $this->client->setClientSecret('YOUR_CLIENT_SECRET');
        $this->client->setRedirectUri(admin_url('admin.php?page=google-auth'));
        $this->client->addScope(Google_Service_Sheets::SPREADSHEETS);
        $this->client->addScope(Google_Service_Drive::DRIVE_FILE);
    }
    public function add_admin_page()
    {
        add_menu_page(
            'Google Sheets Integration',
            'Google Sheets',
            'manage_options',
            'google-auth',
            [$this, 'render_admin_page']
        );
    }

    public function render_admin_page()
    {
        $auth_url = $this->client->createAuthUrl();
        echo '<h1>Google Sheets Integration</h1>';
        echo '<a href="' . esc_url($auth_url) . '" class="button button-primary">Connect Google Account</a>';
    }

    public function handle_oauth()
    {
        if (isset($_GET['code']) && isset($_GET['page']) && $_GET['page'] === 'google-auth') {
            $code = sanitize_text_field($_GET['code']);
            $token = $this->client->fetchAccessTokenWithAuthCode($code);

            if (isset($token['access_token'])) {
                update_option('google_sheets_token', $token);
                wp_redirect(admin_url('admin.php?page=google-auth'));
                exit;
            }
        }
    }
}
