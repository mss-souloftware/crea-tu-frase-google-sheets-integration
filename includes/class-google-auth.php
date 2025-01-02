<?php

class Google_Auth
{
    private $client;

    public function __construct()
    {
        add_action('admin_menu', [$this, 'add_admin_menu']);
        add_action('admin_init', [$this, 'handle_oauth']);
        add_action('admin_post_google_disconnect', [$this, 'handle_disconnect']);
        add_action('admin_post_sync_data_to_google_sheet', [$this, 'sync_data_to_google_sheet']);
        $this->setup_client();
    }

    public function add_admin_menu()
    {
        // Add "Google Sheets" Menu
        add_menu_page(
            'Google Sheets Integration',
            'Google Sheets',
            'manage_options',
            'google-auth',
            [$this, 'render_admin_page']
        );

        // Add "Sync Data" Submenu under "Google Sheets" menu
        add_submenu_page(
            'google-auth', // Parent slug
            'Sync Data', // Page title
            'Sync Data', // Menu title
            'manage_options', // Capability
            'sync-data', // Menu slug
            [$this, 'display_data_from_sheet_and_db'] // Callback function
        );
    }

    private function normalize($data)
    {
        // Assuming the normalization process involves trimming extra spaces and ensuring consistency
        // Example: trim all text fields, convert empty strings to null, and ensure valid numerical fields

        foreach ($data as $key => $value) {
            if (is_string($value)) {
                // Trim extra spaces around string values
                $data[$key] = trim($value);
            }
            // Optionally, handle other types like numbers or booleans here
            if (empty($data[$key]) && $data[$key] !== '0') {
                // Convert empty fields to null (except '0' which is valid)
                $data[$key] = null;
            }
        }
        return $data;
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

    public function render_admin_page()
    {
        $stored_token = get_option('google_sheets_token');

        if ($stored_token) {
            $this->client->setAccessToken($stored_token);

            if ($this->client->isAccessTokenExpired()) {
                $refresh_token = $this->client->getRefreshToken();
                if ($refresh_token) {
                    $new_token = $this->client->fetchAccessTokenWithRefreshToken($refresh_token);

                    if (!isset($new_token['error'])) {
                        // Preserve the refresh token
                        if (isset($stored_token['refresh_token'])) {
                            $new_token['refresh_token'] = $stored_token['refresh_token'];
                        }
                        update_option('google_sheets_token', array_map('sanitize_text_field', $new_token));
                        $this->client->setAccessToken($new_token);
                    } else {
                        $this->handle_token_error();
                        $this->show_reconnect_button();
                        return;
                    }
                } else {
                    $this->handle_token_error();
                    $this->show_reconnect_button();
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

    public function get_access_token()
    {
        // Fetch the stored token from the database (correct field name)
        $accessToken = get_option('google_sheets_token');
        $this->client->setAccessToken($accessToken);

        // Check if the token is expired without refreshing it
        if ($this->client->isAccessTokenExpired()) {
            // Handle expired token error (e.g., show a message to log in again)
            error_log('Access token has expired. Please reauthenticate.');
            wp_die('Google API access token has expired. Please log in again.');
        }

        return $accessToken;
    }




    public function refresh_token_if_needed()
    {
        // Fetch the stored token (correct field name)
        $accessToken = get_option('google_sheets_token'); // Use the correct field name here
        $this->client->setAccessToken($accessToken);

        // If the access token is expired, refresh it
        if ($this->client->isAccessTokenExpired()) {
            error_log('Access token expired, attempting to refresh.');

            // Fetch the refresh token and attempt to refresh the access token
            $refreshToken = $this->client->getRefreshToken();
            if ($refreshToken) {
                $this->client->fetchAccessTokenWithRefreshToken($refreshToken);
                $accessToken = $this->client->getAccessToken();
                update_option('google_sheets_token', $accessToken); // Save the refreshed token
                error_log('Access token refreshed.');
            } else {
                error_log('Refresh token is missing.');
                wp_die('Google API refresh token is not available.');
            }
        } else {
            error_log('Access token is still valid.');
        }
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

        try {
            $service = new Google_Service_Sheets($this->client);

            // Fetch existing data from the sheet to compare
            $response = $service->spreadsheets_values->get($sheet_id, 'Sheet1');
            $sheet_data = $response->getValues();

            // Check if the sheet is empty
            if (empty($sheet_data)) {
                // If the sheet is empty, append all rows from the database
                $body = new Google_Service_Sheets_ValueRange([
                    'values' => $data
                ]);
                $params = ['valueInputOption' => 'RAW'];
                $service->spreadsheets_values->append(
                    $sheet_id,
                    'Sheet1',
                    $body,
                    $params
                );
                echo '<p>All data successfully synced to Google Sheets.</p>';
            } else {
                // If the sheet contains data, compare and update
                $existing_ids = [];
                $rows_to_update = [];
                $rows_to_append = [];

                // Start from the second row (skip header)
                foreach (array_slice($sheet_data, 1) as $index => $sheet_row) {
                    $existing_ids[] = $sheet_row[0]; // Assuming ID is the first column
                }

                // Loop through database rows and compare with sheet data
                foreach ($rows as $row) {
                    $id = $row['id'];
                    $row_data = array_map(function ($value) {
                        return is_null($value) ? '' : $value;
                    }, array_values($row));

                    // If the ID exists in the sheet, update the row
                    if (in_array($id, $existing_ids)) {
                        $row_index = array_search($id, $existing_ids) + 2; // +2 to account for header row and 0-index
                        $rows_to_update[] = ['range' => 'Sheet1!A' . $row_index, 'values' => [$row_data]];
                    } else {
                        // If the ID does not exist in the sheet, append the row
                        $rows_to_append[] = $row_data;
                    }
                }

                // Update rows in the sheet
                foreach ($rows_to_update as $update) {
                    $body = new Google_Service_Sheets_ValueRange([
                        'values' => $update['values']
                    ]);
                    $params = ['valueInputOption' => 'RAW'];
                    $service->spreadsheets_values->update(
                        $sheet_id,
                        $update['range'],
                        $body,
                        $params
                    );
                }

                // Append new rows to the sheet
                if (!empty($rows_to_append)) {
                    $body = new Google_Service_Sheets_ValueRange([
                        'values' => $rows_to_append
                    ]);
                    $params = ['valueInputOption' => 'RAW'];
                    $service->spreadsheets_values->append(
                        $sheet_id,
                        'Sheet1',
                        $body,
                        $params
                    );
                }

                echo '<p>Data successfully updated in Google Sheets.</p>';
            }

        } catch (Exception $e) {
            echo '<p>Error syncing data: ' . esc_html($e->getMessage()) . '</p>';
        }

        wp_die(); // End execution
    }








    private function list_google_sheets()
    {
        // Check if the form is submitted
        if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['google_sheet_id'])) {
            check_admin_referer('save_google_sheet_id', 'google_sheet_id_nonce');

            // Sanitize and save the selected Google Sheet ID
            $selected_id = sanitize_text_field($_POST['google_sheet_id']);
            update_option('selected_google_sheet_id', $selected_id);

            echo '<div class="updated"><p>Google Sheet ID saved successfully!</p></div>';
        }

        try {
            $drive_service = new Google_Service_Drive($this->client);
            $response = $drive_service->files->listFiles([
                'q' => "mimeType='application/vnd.google-apps.spreadsheet'",
                'fields' => 'files(id, name)'
            ]);

            if (!empty($response->files)) {
                echo '<form method="post">';
                wp_nonce_field('save_google_sheet_id', 'google_sheet_id_nonce');
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

    public function display_data_from_sheet_and_db()
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
        $rows = $wpdb->get_results("SELECT * FROM {$table_name}");

        $sheet_data = [];
        $stored_token = get_option('google_sheets_token');
        if ($stored_token) {
            $this->client->setAccessToken($stored_token);
        }

        if ($this->client->getAccessToken()) {
            try {
                $service = new Google_Service_Sheets($this->client);

                // Fetch data from Google Sheet
                $response = $service->spreadsheets_values->get($sheet_id, 'Sheet1');
                $sheet_data = $response->getValues();
            } catch (Exception $e) {
                echo '<p>Error fetching data from Google Sheets: ' . esc_html($e->getMessage()) . '</p>';
            }
        } else {
            echo '<p>Authentication required. Please connect your Google account.</p>';
        }

        // Start output buffering to display data
        ob_start();

        echo '<div class="wrap">';
        echo '<h1>Data from Google Sheet and Database</h1>';

        // Display Google Sheet Data
        echo '<h2>New Data DB</h2>';

        global $wpdb;
        $table_name = $wpdb->prefix . 'chocoletras_plugin';

        // Fetch rows from database
        $db_rows = $wpdb->get_results("SELECT * FROM {$table_name}");

        foreach ($db_rows as $row) {
            echo '<pre>';
            print_r($row);
            echo '</pre>';
        }

        echo '<h2>Google Sheet Data</h2>';
        if (!empty($sheet_data)) {
            // echo '<table border="1" cellspacing="0" cellpadding="5">';
            foreach ($sheet_data as $index => $row) {
                echo '<pre>';
                print_r($sheet_data);
                echo '</pre>';
                // echo '<tr>';
                // foreach ($row as $cell) {
                // echo '<td>' . esc_html($cell) . '</td>';
                // }
                // echo '</tr>';
            }
            // echo '</table>';
        } else {
            echo '<p>No data found in Google Sheet.</p>';
        }

        // Display Database Data
        echo '<h2>Database Data</h2>';
        if (!empty($rows)) {
            // echo '<table border="1" cellspacing="0" cellpadding="5">';
            // Display headers
            // echo '<tr>';
            foreach ($rows as $row) {
                echo '<pre>';
                print_r($row);
                echo '</pre>';
                // echo '<th>' . esc_html($header) . '</th>';
            }
            // echo '</tr>';
            // Display rows
            // foreach ($rows as $row) {
            // echo '<tr>';
            // foreach ($row as $cell) {
            // echo '<td>' . esc_html($cell) . '</td>';
            // }
            // echo '</tr>';
            // }
            // echo '</table>';
        } else {
            echo '<p>No data found in the database.</p>';
        }

        echo '</div>';

        // End output buffering and display content
        echo ob_get_clean();
    }

    private function handle_token_error()
    {
        delete_option('google_sheets_token');
        wp_die('Authentication error. Please reconnect your Google account.');
    }
}