<?php

class Google_Sheets
{
    private $service;

    public function __construct()
    {
        $this->initialize_service();
    }

    private function initialize_service()
    {
        $token = get_option('google_sheets_token');

        if ($token) {
            require_once __DIR__ . '/vendor/autoload.php';

            $client = new Google_Client();
            $client->setAccessToken($token);

            if ($client->isAccessTokenExpired()) {
                $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
                update_option('google_sheets_token', $client->getAccessToken());
            }

            $this->service = new Google_Service_Sheets($client);
        }
    }

    public function add_row_to_sheet($spreadsheetId, $data)
    {
        if (!$this->service) {
            return;
        }

        $range = 'Sheet1!A1';
        $body = new Google_Service_Sheets_ValueRange([
            'values' => [$data]
        ]);

        $this->service->spreadsheets_values->append(
            $spreadsheetId,
            $range,
            $body,
            ['valueInputOption' => 'RAW']
        );
    }
}
