<?php

class Google_Auth
{
    private $client;

    public function __construct()
    {
        add_action('admin_menu', [$this, 'add_admin_page']);
        add_action('admin_init', [$this, 'handle_oauth']);
        add_action('admin_post_google_disconnect', [$this, 'handle_disconnect']);
        add_action('admin_post_sync_data_to_google_sheet', [$this, 'sync_data_to_google_sheet']);
        $this->setup_client();
    }

    private function setup_client()
    {
        require_once __DIR__ . '/lib/vendor/autoload.php';

        $this->client = new Google_Client();
        $this->client->setClientId('');
        $this->client->setClientSecret('');
        $this->client->setRedirectUri(admin_url('admin.php?page=google-auth'));
        $this->client->setAccessType('offline');
        $this->client->setPrompt('consent');
        $this->client->addScope([
            Google_Service_Sheets::SPREADSHEETS,
            Google_Service_Drive::DRIVE,
        ]);
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
        $stored_token = get_option('google_sheets_token');

        // Token validation and refresh
        if ($stored_token) {
            $this->client->setAccessToken($stored_token);

            if ($this->client->isAccessTokenExpired()) {
                $refresh_token = $this->client->getRefreshToken();
                if ($refresh_token) {
                    $new_token = $this->client->fetchAccessTokenWithRefreshToken($refresh_token);
                    if (!isset($new_token['error'])) {
                        update_option('google_sheets_token', $new_token);
                        $this->client->setAccessToken($new_token);
                    } else {
                        $this->handle_token_error();
                        return;
                    }
                } else {
                    $this->handle_token_error();
                    return;
                }
            }
        }

        echo '<h1>Google Sheets Integration</h1>';

        if (!$this->client->getAccessToken()) {
            $auth_url = $this->client->createAuthUrl();
            echo '<a href="' . esc_url($auth_url) . '" class="button button-primary">Connect Google Account</a>';
        } else {
            $this->list_google_sheets();
            echo '<form method="post" action="' . esc_url(admin_url('admin-post.php')) . '">';
            echo '<input type="hidden" name="action" value="sync_data_to_google_sheet">';
            wp_nonce_field('sync_google_sheet_data', 'google_sheet_sync_nonce');
            submit_button('Sync Data to Google Sheets');
            echo '</form>';
            $this->show_disconnect_button();
        }
    }

    public function handle_oauth()
    {
        if (isset($_GET['code']) && isset($_GET['page']) && $_GET['page'] === 'google-auth') {
            $code = sanitize_text_field($_GET['code']);
            $token = $this->client->fetchAccessTokenWithAuthCode($code);

            if (!isset($token['error'])) {
                update_option('google_sheets_token', $token);
                wp_redirect(admin_url('admin.php?page=google-auth'));
                exit;
            } else {
                wp_die('Error during authentication: ' . esc_html($token['error_description'] ?? 'Unknown error'));
            }
        }
    }

    public function handle_disconnect()
    {
        delete_option('google_sheets_token');
        delete_option('selected_google_sheet_id');
        wp_redirect(admin_url('admin.php?page=google-auth'));
        exit;
    }

    private function list_google_sheets()
    {
        try {
            $drive_service = new Google_Service_Drive($this->client);
            $response = $drive_service->files->listFiles([
                'q' => "mimeType='application/vnd.google-apps.spreadsheet'",
                'fields' => 'files(id, name)'
            ]);

            if (!empty($response->files)) {
                echo '<form method="post">';
                echo '<select name="google_sheet_id">';
                foreach ($response->files as $file) {
                    $selected = (get_option('selected_google_sheet_id') == $file->id) ? 'selected' : '';
                    echo '<option value="' . esc_attr($file->id) . '" ' . $selected . '>' . esc_html($file->name) . '</option>';
                }
                echo '</select>';
                submit_button('Save Selection');
                echo '</form>';
            } else {
                echo '<p>No Google Sheets found.</p>';
            }
        } catch (Exception $e) {
            echo '<p>Error retrieving Google Sheets: ' . esc_html($e->getMessage()) . '</p>';
        }
    }

    private function show_disconnect_button()
    {
        echo '<form method="post" action="' . esc_url(admin_url('admin-post.php?action=google_disconnect')) . '">';
        submit_button('Disconnect Google Account', 'primary', 'disconnect');
        echo '</form>';
    }

    public function sync_data_to_google_sheet()
    {
        if (!current_user_can('manage_options')) {
            wp_die('You do not have sufficient permissions to access this page.');
        }

        $sheet_id = get_option('selected_google_sheet_id');
        if (!$sheet_id) {
            wp_die('No Google Sheets selected.');
        }

        global $wpdb;
        $table_name = $wpdb->prefix . 'chocoletras_plugin';
        $rows = $wpdb->get_results("SELECT * FROM {$table_name}", ARRAY_A);

        if (empty($rows)) {
            wp_die('No data to sync.');
        }

        $stored_token = get_option('google_sheets_token');
        if ($stored_token) {
            $this->client->setAccessToken($stored_token);
        }

        if (!$this->client->getAccessToken()) {
            wp_die('Authentication required. Please connect your Google account.');
        }

        // Prepare data
        $header = array_keys(reset($rows)); // Extract column names as the first row
        $data = array_map(function ($row) {
            // Convert associative array to numeric indexed array
            return array_map(function ($value) {
                if (is_array($value)) {
                    return implode(', ', $value); // Flatten arrays to strings
                }
                if (is_null($value)) {
                    return ''; // Convert nulls to empty strings
                }
                return $value;
            }, array_values($row));
        }, $rows);

        array_unshift($data, $header); // Add header as the first row

        // Debugging: Print prepared data
        // echo '<h2>Prepared Data:</h2>';
        // echo '<pre>' . print_r($data, true) . '</pre>';

        // Sync data to Google Sheets
        try {
            $service = new Google_Service_Sheets($this->client);
            $body = new Google_Service_Sheets_ValueRange([
                'values' => $data
            ]);
            $params = ['valueInputOption' => 'RAW'];

            $response = $service->spreadsheets_values->append(
                $sheet_id,
                'Sheet1', // Change to your specific sheet name
                $body,
                $params
            );

            echo '<p>Data successfully synced to Google Sheets.</p>';
            // echo '<pre>' . print_r($response, true) . '</pre>';
        } catch (Exception $e) {
            echo '<p>Error syncing data: ' . esc_html($e->getMessage()) . '</p>';
        }

        wp_die(); // End execution
    }

    private function handle_token_error()
    {
        delete_option('google_sheets_token');
        wp_die('Authentication error. Please reconnect your Google account.');
    }
}
