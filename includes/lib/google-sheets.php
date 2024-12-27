<?php

if ( !defined( 'ABSPATH' ) )
   exit;

include_once ( plugin_dir_path( __FILE__ ) . 'vendor/autoload.php' );

class GSCWOO_googlesheet {

   private $token;
   private $spreadsheet;
   private $worksheet;
   /////=========desktop app
   const clientId_desk = '343006833383-c73msaal124n0psi2fqb5q30u8j2p495.apps.googleusercontent.com';
   const clientSecret_desk = 'k5ThyeLrjEZi0wOhqHhKoWdq';
   //const redirect = 'urn:ietf:wg:oauth:2.0:oob';

   /////=========web app
    const clientId_web = '343006833383-ajjmvck7167u5omiu6kflkmpd7455mo3.apps.googleusercontent.com';
    const clientSecret_web = 'wjSapQopzEaql23EbFNF1Bjk';

   private static $instance;

   public function __construct() {
      
   }

   public static function setInstance( Google_Client $instance = null ) {
      self::$instance = $instance;
   }

   public static function getInstance() {
      if ( is_null( self::$instance ) ) {
         throw new LogicException( "Invalid Client" );
      }

      return self::$instance;
   }

   //constructed on call
   public static function preauth( $access_code ) {
      $newClientSecret = get_option('is_new_client_secret_woogsc');
      $clientId = ($newClientSecret == 1) ? GSCWOO_googlesheet::clientId_web : GSCWOO_googlesheet::clientId_desk;
      $clientSecret = ($newClientSecret == 1) ? GSCWOO_googlesheet::clientSecret_web : GSCWOO_googlesheet::clientSecret_desk;
      

      $client = new Google_Client();
      $client->setClientId($clientId);
      $client->setClientSecret($clientSecret);
      $client->setRedirectUri('https://oauth.gsheetconnector.com');
      $client->setScopes( Google_Service_Sheets::SPREADSHEETS );
      $client->setScopes( Google_Service_Drive::DRIVE_METADATA_READONLY );
      $client->setAccessType( 'offline' );
      $client->fetchAccessTokenWithAuthCode( $access_code );
      $tokenData = $client->getAccessToken();

      GSCWOO_googlesheet::updateToken( $tokenData );
   }

   public static function updateToken($tokenData) {
      $tokenData['expire'] = time() + intval($tokenData['expires_in']);
      try {
         //$tokenJson = json_encode($tokenData);
         //update_option('gfgs_token', $tokenJson);
          //resolved - google sheet permission issues - START
         if(isset($tokenData['scope'])){
            $permission = explode(" ", $tokenData['scope']);
            if((in_array("https://www.googleapis.com/auth/drive.metadata.readonly",$permission)) && (in_array("https://www.googleapis.com/auth/spreadsheets",$permission))) {
               update_option('gs_woo_verify', 'valid');
            } else {
                 // update_option('gs_verify', 'Something went wrong! It looks you have not given the permission of Google Drive and Google Sheets from your google account.Please Deactivate Auth and Re-Authenticate again with the permissions.');
                 update_option('gs_woo_verify','invalid-auth');
            }
         }
         $tokenJson = json_encode($tokenData);
         update_option('gs_woo_token', $tokenJson);
         //resolved - google sheet permission issues - END

      } catch (Exception $e) {
         wc_gsheetconnector_utility::gs_debug_log("Token write fail! - " . $e->getMessage());
      }
   }

   public function auth() {
      $tokenData = json_decode( get_option( 'gs_woo_token' ), true );
      if ( !isset( $tokenData['refresh_token'] ) || empty( $tokenData['refresh_token'] ) ) {
         throw new LogicException( "Auth, Invalid OAuth2 access token" );
         exit();
      }

      try {
         $newClientSecret = get_option('is_new_client_secret_woogsc');
         $clientId = ($newClientSecret == 1) ? GSCWOO_googlesheet::clientId_web : GSCWOO_googlesheet::clientId_desk;
         $clientSecret = ($newClientSecret == 1) ? GSCWOO_googlesheet::clientSecret_web : GSCWOO_googlesheet::clientSecret_desk;
         $client = new Google_Client();
         $client->setClientId($clientId);
         $client->setClientSecret($clientSecret);

         $client->setScopes( Google_Service_Sheets::SPREADSHEETS );
         $client->setScopes( Google_Service_Drive::DRIVE_METADATA_READONLY );
         $client->refreshToken( $tokenData['refresh_token'] );
         $client->setAccessType( 'offline' );
         GSCWOO_googlesheet::updateToken( $tokenData );

         self::setInstance( $client );
      } catch ( Exception $e ) {
         throw new LogicException( "Auth, Error fetching OAuth2 access token, message: " . $e->getMessage() );
         exit();
      }
   }

   public function setSpreadsheetId( $id ) {
      $this->spreadsheet = $id;
   }

   public function getSpreadsheetId() {

      return $this->spreadsheet;
   }

   public function setWorkTabId( $id ) {
      $this->worksheet = $id;
   }

   public function getWorkTabId() {
      return $this->worksheet;
   }
   
   public function add_row($data_value, $gscwoo_operation, $gscwoo_sheetname, $gscwoo_order_id) {         
      try{
         $client = self::getInstance();
         $service = new Google_Service_Sheets($client);
         $spreadsheetId = $this->getSpreadsheetId();
         $work_sheets = $service->spreadsheets->get($spreadsheetId);
         /* new Update  START */
         if (!empty($work_sheets) && !empty($data_value)) {
                  foreach ($work_sheets as $sheet) {
                     $properties = $sheet->getProperties();
                     $p_title = $properties->getSheetId();
                     $w_title = $this->getWorkTabId();
                     if ($p_title == $w_title) {
                        $w_title = $properties->getTitle();
                        $worksheetCell = $service->spreadsheets_values->get($spreadsheetId, $w_title . "!1:1");
                        $insert_data = array();
                        if (isset($worksheetCell->values[0])) {
                           $insert_data_index = 0;
                           foreach ($worksheetCell->values[0] as $k => $name) {
                              if ($insert_data_index == 0) {
                                 if (isset($data_value[$name]) && $data_value[$name] != '') {
                                    $insert_data[] = $data_value[$name];
                                 } else {
                                    $insert_data[] = '';
                                 }
                              } else {
                                 if (isset($data_value[$name]) && $data_value[$name] != '') {
                                    $insert_data[] = $data_value[$name];
                                 } else {
                                    $insert_data[] = '';
                                 }
                              }
                              $insert_data_index++;
                           }
                        }
                     $tab_name = $w_title;
                     $full_range = $tab_name."!A1:Z";
                     $response   = $service->spreadsheets_values->get( $spreadsheetId, $full_range );
                     $get_values = $response->getValues();
                     
                     if( $get_values) {
                        $row  = count( $get_values ) + 1;
                     }
                     else {
                        $row = 1;
                     }
                     /* Get the range of sheet - START */
                     if($gscwoo_operation == "update"){
                        $gscwoo_sheet  = "'".$gscwoo_sheetname."'!A:A";
                        $gscwoo_allentry = $service->spreadsheets_values->get($spreadsheetId, $gscwoo_sheet);
                        $gscwoo_data   = $gscwoo_allentry->getValues();
                        $counter = 1;
                        foreach ($gscwoo_data as $key => $value) {
                           if($value[0] == $gscwoo_order_id){
                              $gscwoo_num = $counter;
                              break;
                           }
                           $counter++;
                        }
                        $range = $tab_name.'!A'.$gscwoo_num; 
                     }
                     /* Get the range of sheet - END */

                     /* Add and Delete Row - START */
                     if($gscwoo_operation == "add_delete"){
                        $gscwoo_sheet  = "'".$gscwoo_sheetname."'!A:A";
                        $gscwoo_allentry = $service->spreadsheets_values->get($spreadsheetId, $gscwoo_sheet);
                        $gscwoo_data   = $gscwoo_allentry->getValues();
                        $counter = 1;
                        foreach ($gscwoo_data as $key => $value) {
                           if($value[0] == $gscwoo_order_id){
                              $gscwoo_num = $counter;
                              break;
                           }
                           $counter++;
                        }
                        $range = $tab_name.'!A'.$gscwoo_num; 
                     }
                     /* Add and Delete Row - END */

                     if($gscwoo_operation == "insert")
                           $range = $tab_name."!A".$row.":Z";
                     
                        $range_new = $w_title;
                        // Create the value range Object
                        $valueRange = new Google_Service_Sheets_ValueRange();
                        // You need to specify the values you insert
                        $valueRange->setValues(["values" => $insert_data]);
                        // Add two values
                        // Then you need to add some configuration
                        $conf = ["valueInputOption" => "USER_ENTERED", "insertDataOption" => "INSERT_ROWS"];
                       $conf = ["valueInputOption" => "USER_ENTERED"];
                        // append the spreadsheet
                       if($gscwoo_operation == "update")
                           $result = $service->spreadsheets_values->update($spreadsheetId, $range, $valueRange, $conf);

                       if($gscwoo_operation == "add_delete")
                       {
                        // Code Pending From - 09-06-2021
                        // echo "=============== add update =================";
                        // $conf = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest(array(
                        //       'requests' => array(
                        //          'deleteDimension' => array(
                        //             'range' => array(
                        //               'dimension'  => 'ROWS',
                        //               'sheetId'    => $spreadsheetId,
                        //               'startIndex'    => $gscwoo_num,
                        //               'endIndex'   => $gscwoo_num+1
                        //             )
                        //          )
                        //       )));
                        //    $result = $service->spreadsheets_values->batchUpdate($spreadsheetId, $range, $valueRange, $conf);
                        // Code Pending From - 09-06-2021
                       }
                       
                        
                       if($gscwoo_operation == "insert")
                           $result = $service->spreadsheets_values->append($spreadsheetId, $range, $valueRange, $conf);

                     }
                  }
               }
            } catch (Exception $e) {
                  return null;
                  exit();
               }
         }         

   public function add_multiple_row( $data ) {
      try {
         $client = self::getInstance();
         $service = new Google_Service_Sheets( $client );
         $spreadsheetId = $this->getSpreadsheetId();
         $work_sheets = $service->spreadsheets->get( $spreadsheetId );

         if ( !empty( $work_sheets ) && !empty( $data ) ) {
            foreach ( $work_sheets as $sheet ) {
               $properties = $sheet->getProperties();
               $sheet_id = $properties->getSheetId();

               $worksheet_id = $this->getWorkTabId();

               if ( $sheet_id == $worksheet_id ) {
                  $worksheet_id = $properties->getTitle();
                  $worksheetCell = $service->spreadsheets_values->get( $spreadsheetId, $worksheet_id . "!1:1" );
                  $insert_data = array();
                  $final_data = array();
                  if ( isset( $worksheetCell->values[0] ) ) {
                     foreach ( $data as $key => $value ) {
                        foreach ( $worksheetCell->values[0] as $k => $name ) {
                           if ( isset( $value[$name] ) && $value[$name] != '' ) {
                              $insert_data[] = $value[$name];
                           } else {
                              $insert_data[] = '';
                           }
                        }
                        $final_data[] = $insert_data;
                        unset( $insert_data );
                     }
                  }

                  $range_new = $worksheet_id;

                  $sheet_values = $final_data;

                  if ( !empty( $sheet_values ) ) {
                     $requestBody = new Google_Service_Sheets_ValueRange( [
                        'values' => $sheet_values
                             ] );

                     $params = [
                        'valueInputOption' => 'USER_ENTERED'
                     ];
                     $response = $service->spreadsheets_values->append( $spreadsheetId, $range_new, $requestBody, $params );
                  }
               }
            }
         }
      } catch ( Exception $e ) {
         return null;
         exit();
      }
   }

   //get all the spreadsheets
   public function get_spreadsheets() {
      $all_sheets = array();
      try {
         $client = self::getInstance();

         $service = new Google_Service_Drive( $client );

         $optParams = array(
            'q' => "mimeType='application/vnd.google-apps.spreadsheet'"
         );
         $results = $service->files->listFiles( $optParams );
         foreach ( $results->files as $spreadsheet ) {
            if ( isset( $spreadsheet['kind'] ) && $spreadsheet['kind'] == 'drive#file' ) {
               $all_sheets[] = array(
                  'id' => $spreadsheet['id'],
                  'title' => $spreadsheet['name'],
               );
            }
         }
      } catch ( Exception $e ) {
         return null;
         exit();
      }
      return $all_sheets;
   }

   //get worksheets title
   public function get_worktabs( $spreadsheet_id ) {
      $work_tabs_list = array();
      try {
         $client = self::getInstance();
         $service = new Google_Service_Sheets( $client );
         $work_sheets = $service->spreadsheets->get( $spreadsheet_id );


         foreach ( $work_sheets as $sheet ) {
            $properties = $sheet->getProperties();
            $work_tabs_list[] = array(
               'id' => $properties->getSheetId(),
               'title' => $properties->getTitle(),
            );
         }
      } catch ( Exception $e ) {
         return null;
         exit();
      }

      return $work_tabs_list;
   }

   /**
    * Function - Adding custom column header to the sheet
    * @param string $sheet_name
    * @param string $tab_name
    * @param array $gs_map_tags 
    * @since 1.0
    */
   public function add_header( $sheetname, $tabname, $final_header_array, $old_header ) {
      $client = self::getInstance();
      $service = new Google_Service_Sheets( $client );
      $spreadsheetId = $this->getSpreadsheetId();
      $work_sheets = $service->spreadsheets->get( $spreadsheetId );

      $field_tag_array[] = '';
      if ( !empty( $work_sheets ) ) {
         foreach ( $work_sheets as $sheet ) {

            $properties = $sheet->getProperties();
            $sheet_id = $properties->getSheetId();
            $worksheet_id = $this->getWorkTabId();
            if ( $sheet_id == $worksheet_id ) {
               $worksheet_title = $properties->getTitle();
               $field_tag = isset( $_POST['gf-custom-ck'] ) ? sanitize_text_field($_POST['gf-custom-ck']) : array();
               $field_tag_key = isset( $_POST['gf-custom-header-key'] ) ? sanitize_text_field($_POST['gf-custom-header-key']) : "";
               $field_tag_placeholder = isset( $_POST['gf-custom-header-placeholder'] ) ? sanitize_text_field($_POST['gf-custom-header-placeholder']) : "";
               $field_tag_column = isset( $_POST['gf-custom-header'] ) ? sanitize_text_field($_POST['gf-custom-header']) : "";
               if ( !empty( $field_tag ) ) {
                  foreach ( $field_tag as $key => $value ) {
                     $gf_key = $field_tag_key[$key];
                     $gf_val = (!empty( $field_tag_column[$key] ) ) ? $field_tag_column[$key] : $field_tag_placeholder[$key];
                     if ( $gf_val !== "" ) {
                        $field_tag_array[$gf_key] = $gf_val;
                        $gravityform_tags[] = $gf_val;
                     }
                  }
               }
               $range = $worksheet_title . '!1:1';

               $values = array( array_values( array_filter( $field_tag_array ) ) );


               $count_old_header = count( $old_header );
               $count_new_header = count( $final_header_array );
               $data_values = array();

// If old header count is greater than new header count than empty the header
               if ( $count_old_header !== 0 && $count_old_header > $count_new_header ) {
                  for ( $i = 0; $i <= $count_old_header; $i++ ) {
                     $column_name = isset( $final_header_array[$i] ) ? $final_header_array[$i] : "";
                     if ( $column_name !== "" ) {
                        $data_values[] = $column_name;
                     } else {
                        $data_values[] = "";
                     }
                  }
               } else {

                  foreach ( $final_header_array as $column_name ) {
                     $data_values[] = $column_name;
                  }
               }

               $values = array( $data_values );


               $requestBody = new Google_Service_Sheets_ValueRange( [
                  'values' => $values
               ] );

               $params = [
                  'valueInputOption' => 'RAW'
               ];
               $response = $service->spreadsheets_values->update( $spreadsheetId, $range, $requestBody, $params );
            }
         }
      }
   }
   
   public function perform_sheet_tab_updates( $spreadsheet_id, $request_array ) {
		$gscwoo_client	 = self::getInstance();
		$gscwoo_service	 = new Google_Service_Sheets( $gscwoo_client );
		$update_request = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest( array( 'requests' => $request_array ) );
		$gscwoo_response = $gscwoo_service->spreadsheets->batchUpdate($spreadsheet_id, $update_request);
   }
   
   
   /**
    * Create, Insert and Update Tabs and Headers
    * @param int $selected_sheet_id
    * @param string $gscwoo_spreadsheetName
    * @param array $order_states
    * @since 1.0
    */
   public function ciu_tabs_and_headers( $selected_sheet_id, $gscwoo_spreadsheetName, $order_states ) {
        // Get header list
		
		$gs_service = new wc_gsheetconnector_Service();
		
		$header_list = array_keys( $gs_service->sheet_headers );
		$gscwoo_sheetnames = array_keys ( $gs_service->status_and_sheets );
		
		$gscwoo_remove_sheet = array();
		
		echo "<pre>"; print_r($gscwoo_remove_sheet); echo "</pre>"; exit;
		
		/*//if( isset( $order_states['wcpendingorder'] ) ) { 
		//if( false !== array_search( 'wcpendingorder', $order_states ) ) {
		if( in_array( 'wcpendingorder', $order_states ) ) {
			$gscwoo_pendingorder  = 1;
		}else{ 
			$gscwoo_pendingorder  = 0; 
			$gscwoo_remove_sheet[] = 'Pending Orders';
		}
		
		if( in_array( 'wcprocessingorder', $order_states ) ) { 
			$gscwoo_processingorder  = 1;
		}else{ 
			$gscwoo_processingorder  = 0; 
			$gscwoo_remove_sheet[] = 'Processing Orders';
		}
		
		if( in_array( 'wconholdorder', $order_states ) ) { 
			$gscwoo_onholdorder  = 1;
		}else{ 
			$gscwoo_onholdorder  = 0; 
			$gscwoo_remove_sheet[] = 'On Hold Orders';
		}
		
		if( in_array( 'wccompletedorder', $order_states ) ) { 
			$gscwoo_completedorders  = 1;
		}else{ 
			$gscwoo_completedorders  = 0; 
			$gscwoo_remove_sheet[] = 'Completed Orders';
		}
		
		if( in_array( 'wccancelledorder', $order_states ) ) { 
			$gscwoo_cancelledorders  = 1;
		}else{ 
			$gscwoo_cancelledorders  = 0; 
			$gscwoo_remove_sheet[] = 'Cancelled Orders';
		}
		
		if( in_array( 'wcrefundedorder', $order_states ) ) { 
			$gscwoo_refundedorders  = 1;
		}else{ 
			$gscwoo_refundedorders  = 0; 
			$gscwoo_remove_sheet[] = 'Refunded Orders';
		}
		
		if( in_array( 'wcfailedorder', $order_states ) ) { 
			$gscwoo_failedorders  = 1;
		}else{ 
			$gscwoo_failedorders  = 0; 
			$gscwoo_remove_sheet[] = 'Failed Orders';
		}
		
		if( in_array( 'wctrashorder', $order_states ) ) { 
			$gscwoo_trashorders  = 1;
		}else{ 
			$gscwoo_trashorders  = 0; 
			$gscwoo_remove_sheet[] = 'Trash Orders';
		}
		
		$gscwoo_order_array = array( $gscwoo_pendingorder, $gscwoo_processingorder, $gscwoo_onholdorder, $gscwoo_completedorders, $gscwoo_cancelledorders ,$gscwoo_refundedorders, $gscwoo_failedorders, $gscwoo_trashorders );
			
		$gscwoo_client	 = self::getInstance();
		$gscwoo_service	 = new Google_Service_Sheets( $gscwoo_client );
		$gscwoo_spreadsheetId = $this->getSpreadsheetId();
		$gscwoo_response = $gscwoo_service->spreadsheets->get( $gscwoo_spreadsheetId );
		foreach ( $gscwoo_response->getSheets() as $gscwoo ) {
			$existing_sheet_tabs[ $gscwoo[ 'properties' ][ 'sheetId' ] ] = $gscwoo[ 'properties' ][ 'title' ];
		}
		
		for ( $i = 0; $i < count( $gscwoo_sheetnames ); $i ++ ) {
			if ( in_array( $gscwoo_sheetnames[ $i ], $existing_sheet_tabs ) ) {
			$gscwoo_order_array[ $i ] = 0;
			} else {
			if ( isset( $order_states[ $gscwoo_sheetnames[ $i ] ] ) && ! in_array( $gscwoo_sheetnames[ $i ], $existing_sheet_tabs ) )
				$gscwoo_order_array[ $i ] = 1;
			}
		}
		
		// If have to create new tab
		for( $i = 0; $i<count( $gscwoo_order_array ); $i++ ) { 
			if( $i == 0 ){ $gscwoo_sheetname = 'Pending Orders'; }
			if( $i == 1 ){ $gscwoo_sheetname = 'Processing Orders'; } 
			if( $i == 2 ){ $gscwoo_sheetname = 'On Hold Orders'; }
			if( $i == 3 ){ $gscwoo_sheetname = 'Completed Orders'; }
			if( $i == 4 ){ $gscwoo_sheetname = 'Cancelled Orders'; } 
			if( $i == 5 ){ $gscwoo_sheetname = 'Refunded Orders'; }
			if( $i == 6 ){ $gscwoo_sheetname = 'Failed Orders'; } 
			if( $i == 7 ){ $gscwoo_sheetname = 'Trash Orders'; }

			if( $gscwoo_order_array[$i] == '1' ){ 			
			$gscwoo_body = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest( array(
				'requests' => array(
				'addSheet' => array(
					'properties' => array(
					'title' => $gscwoo_sheetname
					)
				)
				)
			) );
			$gscwoo_response = $gscwoo_service->spreadsheets->batchUpdate($gscwoo_spreadsheetId,$gscwoo_body);
			$gscwoo_range = trim($gscwoo_sheetname).'!A1';
			$gscwoo_requestBody = new Google_Service_Sheets_ValueRange(array(
				'values' => array( stripslashes_deep($header_list) )
			));

			$gscwoo_params = array( 'valueInputOption' => 'USER_ENTERED' ); 

			$gscwoo_response = $gscwoo_service->spreadsheets_values->update($gscwoo_spreadsheetId, $gscwoo_range, $gscwoo_requestBody, $gscwoo_params);
			}
		}
		
		// If already tab exist then just save headers
		for ( $i = 0; $i < count( $gscwoo_sheetnames ); $i ++  ) {

			if ( in_array( $gscwoo_sheetnames[ $i ], $existing_sheet_tabs ) && $gscwoo_order_array[ $i ] == '0' ) {
			$gscwoo_range		 = trim( $gscwoo_sheetnames[ $i ] ) . '!A1';
			$gscwoo_requestBody	 = new Google_Service_Sheets_ValueRange( array(
				'values' => array( stripslashes_deep($header_list) )
			) );
			$gscwoo_params		 = array( 'valueInputOption' => 'USER_ENTERED' );
			$gscwoo_response = $gscwoo_service->spreadsheets_values->update($gscwoo_spreadsheetId, $gscwoo_range, $gscwoo_requestBody, $gscwoo_params);
			}
		}*/
	}
	
   public function insert_data_into_sheet( $gscwoo_operation, $gscwoo_order_id, $selected_sheet_id, $gscwoo_sheetname) {

      $tabId = $this->getTabId($selected_sheet_id, $gscwoo_sheetname);
      $doc = new GSCWOO_googlesheet();
      $doc->auth();
      $doc->setSpreadsheetId($selected_sheet_id);
      $doc->setWorkTabId($tabId);
      $gscwoo_value_array = wc_gsheetconnector_utility::make_values_array( $gscwoo_operation, $gscwoo_order_id );
      $doc->add_row($gscwoo_value_array, $gscwoo_operation, $gscwoo_sheetname, $gscwoo_order_id);
    }
	
	public function getTabId($selected_sheet_id, $gscwoo_sheetname){
		$tabsArr = $this->get_worktabs( $selected_sheet_id );
		foreach ($tabsArr as $key => $value) {
			if($value["title"] == $gscwoo_sheetname)
			$tabId = $value["id"];
		}
		return $tabId;
	}
	
	
	/**************************************************************
	** FUNCTIONS BY RASHID **
	**************************************************************/
	
	public function get_sheet_tabs($spreadsheet_id) {
		$tabs = $this->get_worktabs($spreadsheet_id);
		$tabs = wp_list_pluck( $tabs, "title", "id" );
		return $tabs;
	}
	
	public function get_sheet_name( $spreadsheet_id, $tab_id ) {
		
		$all_sheet_data = get_option( 'gs_woo_sheetId' );
		
		$tab_name = "";
		foreach( $all_sheet_data as $spreadsheet ) {
			
			if( $spreadsheet['id'] == $spreadsheet_id ) {
				$tabs = $spreadsheet['tabId'];
				
				foreach( $tabs as $name => $id ) {
					if( $id == $tab_id ) {
						$tab_name = $name;
					}
				}
			}
		}
		
		$tab_name = apply_filters( "gcwoo_filter_tab_name", $tab_name, $spreadsheet_id, $tab_id );		
		return $tab_name;
	}
	
	public function get_spreadsheet_name( $spreadsheet_id ) {
		
		$all_sheet_data = get_option( 'gs_woo_sheetId' );
		
		$spreadsheetName = "";
		foreach( $all_sheet_data as $spreadsheet_name => $spreadsheet ) {
			
			if( $spreadsheet['id'] == $spreadsheet_id ) {
				$spreadsheetName = $spreadsheet_name;
			}
		}
		
		$spreadsheetName = apply_filters( "gcwoo_filter_spreasheet_name", $spreadsheetName, $spreadsheet_id );
		
		return $spreadsheetName;
	}
	
	public function add_bulk_rows_to_sheet( $spreadsheet_id, $tab_name, $row_data_arrays, $processed_entries = array() ) {
		
		if( ! $row_data_arrays ) {
			return;
		}
		
		$client = self::getInstance();	
		
		if( ! $client ) {
			return false;
		}
		
		try {		
			
			$service = new Google_Service_Sheets($client);
			$full_range = $tab_name."!A1:Z";
			$response   = $service->spreadsheets_values->get( $spreadsheet_id, $full_range );
			$get_values = $response->getValues();
			
			if( $get_values) {
				$row  = count( $get_values ) + 1;
			}
			else {
				$row = 1;
			}
			
			$total_row_data = count( $row_data_arrays );
			$start_index = $row;
			$end_index = $row + $total_row_data;
			
			foreach( $row_data_arrays as &$row_data ) {				
				ksort($row_data);
				
				foreach( $row_data as &$data ) {
					$data = str_replace( "{row}", $row, $data );
				}
				$row++;
			}
			
			$range = $tab_name."!A".$start_index.":A";
			$valueRange = new Google_Service_Sheets_ValueRange();

			$valueRange->setValues($row_data_arrays);

			// $range = 'Sheet1!A1:A';
			$conf = ["valueInputOption" => "USER_ENTERED"];
			$service->spreadsheets_values->append($spreadsheet_id, $range, $valueRange, $conf);
			
			do_action( "gcwoo_after_bulk_entries_added", $row_data_arrays, $processed_entries );
			return true;
		} 
		catch (Exception $e) {
			return false;
		}
	}
	
	public function remove_row_by_order_id( $spreadsheet_id, $tab_name, $order_id, $order_id_index ) {
		
		$client = self::getInstance();	
			
		if( ! $client ) {
			return false;
		}
		
		try {
			$tab_id = $this->getTabId( $spreadsheet_id, $tab_name );
			$service = new Google_Service_Sheets($client);
			$full_range = $tab_name."!A1:Z";
			$response   = $service->spreadsheets_values->get( $spreadsheet_id, $full_range );
			$get_values = $response->getValues();
			
			$order_ids = wp_list_pluck( $get_values, $order_id_index );
			
			/*$order_rows_indexes = array_keys( $order_ids, $order_id );		
			$requests = array();
			$tab_id = $this->getTabId( $spreadsheet_id, $tab_name );
			foreach( $order_rows_indexes as $index ) {
				
				$requests[] = array(
					'deleteDimension' => array(
						'range' => array(
							'dimension'  => 'ROWS',
							'sheetId'    => $tab_id,
							'startIndex'    => $index,
							'endIndex'   => $index+1
						),
					)
				);
				
			}
			
			$conf = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest(
				array(
					'requests' => $requests
				)
			);*/
			
			$index = array_search($order_id, $order_ids);
			
			if( $index != false ) {
				
				$conf = array(
					'requests' => array(
						'deleteDimension' => array(
							'range' => array(
								'dimension'  => 'ROWS',
								'sheetId'    => $tab_id,
								'startIndex'    => $index,
								'endIndex'   => $index+1
							)
						)
					)
				);
				
				$conf = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest($conf);
				
				$result = $service->spreadsheets->batchUpdate($spreadsheet_id, $conf);	
			}
			
		}
		catch (Exception $e) {
			echo $e->getMessage();
			return false;
		}
	}
	
	public function update_row_by_order_id( $spreadsheet_id, $tab_name, $row_data, $order_id, $order_id_index ) {
		
		$client = self::getInstance();	
			
		if( ! $client ) {
			return false;
		}
		
		try {
			$tab_id = $this->getTabId( $spreadsheet_id, $tab_name );
			$service = new Google_Service_Sheets($client);
			$full_range = $tab_name."!A1:Z";
			$response   = $service->spreadsheets_values->get( $spreadsheet_id, $full_range );
			$get_values = $response->getValues();
			
			$order_ids = wp_list_pluck( $get_values, $order_id_index );
			$row = array_search($order_id, $order_ids);
			
			foreach($row_data as &$data) {
				$data = str_replace( "{row}", $row, $data );
			}
			
			if( $row === false ) {
				if( $get_values) {
					$row  = count( $get_values ) + 1;
				}
				else {
					$row = 1;
				}
				
				$range = $tab_name."!A".$row.":Z";
				$valueRange = new Google_Service_Sheets_ValueRange();
				$valueRange->setValues(["values" => $row_data]);
				$conf = ["valueInputOption" => "USER_ENTERED", "insertDataOption" => "INSERT_ROWS"];
				$result = $service->spreadsheets_values->append($spreadsheet_id, $range, $valueRange, $conf);	
			}
			else {
				$row = $row+1;
				$range = $tab_name."!A".$row.":".$row;
				$valueRange = new Google_Service_Sheets_ValueRange();
				$valueRange->setValues(["values" => $row_data]);
				$conf = ["valueInputOption" => "USER_ENTERED"];
				$result = $service->spreadsheets_values->update($spreadsheet_id, $range, $valueRange, $conf);
			}
			
			/*$conf = array(
				'requests' => array(
					'deleteDimension' => array(
						'range' => array(
							'dimension'  => 'ROWS',
							'sheetId'    => $tab_id,
							'startIndex'    => $index,
							'endIndex'   => $index+1
						)
					)
				)
			);
			
			$conf = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest($conf);
			
			$result = $service->spreadsheets->batchUpdate($spreadsheet_id, $conf);	*/
			
		}
		catch (Exception $e) {
			echo '*****'.$e->getMessage();
			return false;
		}
	}
	
	public function add_row_to_sheet( $spreadsheet_id, $tab_name, $row_data, $order, $is_header = false ) {
		
		if( ! $row_data ) {
			return;
		}
		
		ksort($row_data);
		
		try {			
			$client = self::getInstance();	
			
			if( ! $client ) {
				return false;
			}
					
			$service = new Google_Service_Sheets($client);
			
			
			$full_range = $tab_name."!A1:Z";
			$response   = $service->spreadsheets_values->get( $spreadsheet_id, $full_range );
			$get_values = $response->getValues();
			
			if( $get_values) {
				$row  = count( $get_values ) + 1;
			}
			else {
				$row = 1;
			}
			
			foreach($row_data as &$data) {
				$data = str_replace( "{row}", $row, $data );
			}
			
			
			if( $is_header ) { 
				$range = $tab_name . '!1:1';
				$valueRange = new Google_Service_Sheets_ValueRange();
				$valueRange->setValues(["values" => $row_data]);
				$conf = ["valueInputOption" => "RAW"];
				$result = $service->spreadsheets_values->update($spreadsheet_id, $range, $valueRange, $conf);	
				do_action( "gcwoo_header_updated", $row_data );
			}
			else {
				$range = $tab_name."!A".$row.":Z";
				$valueRange = new Google_Service_Sheets_ValueRange();
				$valueRange->setValues(["values" => $row_data]);
				$conf = ["valueInputOption" => "USER_ENTERED", "insertDataOption" => "INSERT_ROWS"];
				$result = $service->spreadsheets_values->append($spreadsheet_id, $range, $valueRange, $conf);					
				do_action( "gcwoo_entry_added", $row_data, $order );
			}
			return true;
		} 
		catch (Exception $e) {
			// echo $e->getMessage();
			return false;
		}
	}
	
	public function get_header_row( $spreadsheet_id, $tab_id ) {
		
		$header_cells = array();
		try {
		
			$client = $this->getInstance();			
			
			if( ! $client ) {
				return false;
			}			
			
			$service = new Google_Service_Sheets($client);
			
			$work_sheets = $service->spreadsheets->get($spreadsheet_id);
			
			if( $work_sheets ) {
				
				foreach ($work_sheets as $sheet) {
				
					$properties = $sheet->getProperties();
					$work_sheet_id = $properties->getTitle();
					
					if( $work_sheet_id == $tab_id ) {
						
						$tab_title = $properties->getTitle();
						$header_row = $service->spreadsheets_values->get($spreadsheet_id, $tab_title . "!1:1");
						
						$header_row_values = $header_row->getValues();
						
						if( isset( $header_row_values[0] ) && $header_row_values[0] ) {
							$header_cells = $header_row_values[0];
						}		
					}
				}
			}
		}
		catch (Exception $e) {
			$header_cells = array();
		}
		
		$header_cells = apply_filters( "gcwoo_fetched_header_cells", $header_cells, $spreadsheet_id, $tab_id );
		
		return $header_cells;
	}
	
	public function sort_sheet_by_column( $spreadsheet_id, $tab_id, $column_index, $sort_order = "ASCENDING" ) {
		
		try {
			if( $column_index !== false && is_numeric($column_index) ) {			
				$client = $this->getInstance();
				
				if( ! $client ) {
					return false;
				}
				
				$service = new Google_Service_Sheets($client);
				
				$args = array(
					"sortRange" => array(
						'range' => array(
							'sheetId' => $tab_id,
							'startRowIndex' => 1,
							'startColumnIndex' => 0,
						),
						'sortSpecs' => array(
							array(
								'sortOrder' => $sort_order,
								'dimensionIndex' => $column_index,
							),							
						),
					)
				);
				
				$google_service_sheet_request = new Google_Service_Sheets_Request( $args );				
				$request = array( $google_service_sheet_request );				
				$args = array( "requests" => $request );
				$batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest( $args );
				$result = $service->spreadsheets->batchUpdate($spreadsheet_id, $batchUpdateRequest);
				return true;
			}
		}
		
		catch (Exception $e) {
			return false;
		}
	}
	
	function hex_color_to_google_rgb($hex) {
	
		$rgb_return = array();
		
		$hex      = str_replace('#', '', $hex);
		$length   = strlen($hex);
		$rgb['red'] = hexdec($length == 6 ? substr($hex, 0, 2) : ($length == 3 ? str_repeat(substr($hex, 0, 1), 2) : 0));
		$rgb['green'] = hexdec($length == 6 ? substr($hex, 2, 2) : ($length == 3 ? str_repeat(substr($hex, 1, 1), 2) : 0));
		$rgb['blue'] = hexdec($length == 6 ? substr($hex, 4, 2) : ($length == 3 ? str_repeat(substr($hex, 2, 1), 2) : 0));
		
		foreach( $rgb as $key => $clr ) {
			$rgb_return[$key] = $clr / 255;
		}
		return $rgb_return;
	}
	
	public function freeze_row( $spreadsheet_id, $tab_id, $number_of_rows = 1 ) {
	
		$number_of_rows = apply_filters( "gsheet_default_frozen_rows", $number_of_rows );
		
		try {
			$client = $this->getInstance();	
			
			if( ! $client ) {
				return false;
			}
			
			$service = new Google_Service_Sheets($client);
			$args = array(
				"updateSheetProperties" => array(
					'fields' => 'gridProperties.frozenRowCount',
					'properties' => [
						'sheetId' => $tab_id,
						'gridProperties' => array(
							'frozenRowCount' => $number_of_rows
						),
					],
				)
			);
			
			$google_service_sheet_request = new Google_Service_Sheets_Request( $args );				
			$request = array( $google_service_sheet_request );				
			$args = array( "requests" => $request );
			$batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest( $args );
			$result = $service->spreadsheets->batchUpdate($spreadsheet_id, $batchUpdateRequest);
			
			return true;
		}
		catch (Exception $e) {
			return false;
		}
	}

	public function set_alternate_colors( $spreadsheet_id, $tab_id, $headerColor, $oddColor, $evenColor ) {
	
		try {
			$client = $this->getInstance();	
			
			if( ! $client ) {
				return false;
			}
			
			$service = new Google_Service_Sheets($client);
			$work_sheets = $service->spreadsheets->get($spreadsheet_id);
			
			$range = array( 'sheetId' => $tab_id );
			$args = array();
			
			$range_exist = false;
			
			$rowProperties = array();
			$rowProperties["headerColor"] = $headerColor ? $this->hex_color_to_google_rgb($headerColor) : $this->hex_color_to_google_rgb("#ffffff");
			$rowProperties["firstBandColor"] = $oddColor ? $this->hex_color_to_google_rgb($oddColor) : $this->hex_color_to_google_rgb("#ffffff");
			$rowProperties["secondBandColor"] = $evenColor ? $this->hex_color_to_google_rgb($evenColor) : $this->hex_color_to_google_rgb("#ffffff");
			
			$banded_range_id = 100;
			if( $tab_id != 0 ) {
				$generate_banded_range_id = substr($tab_id, 0, 4);
				$banded_range_id = $generate_banded_range_id;
			}
			
			$banding_request = array(	
				"bandedRange" => array(
					"bandedRangeId" => $banded_range_id,
					"range" => $range,
					"rowProperties" => $rowProperties,
				)
			);	
			
			
			foreach ($work_sheets as $sheet) {
				$properties = $sheet->getProperties();			
				if( $properties->sheetId == $tab_id ) {				
					$bandedRanges = $sheet->getBandedRanges();
					foreach( $bandedRanges as $bandedRange	 ) {					
						if( $bandedRange->bandedRangeId == $banded_range_id ) {
							$range_exist = true;
						}
					}
				}
			}
			
			if( $range_exist ) {
				$args['updateBanding'] = $banding_request;
				$args['updateBanding']['fields'] = "*";
			}
			else {
				$args['addBanding'] = $banding_request;
			}
			
			$banding_request = new Google_Service_Sheets_Request( $args );	
			$request = array( $banding_request );			
			
			$args = array( "requests" => $request );
			$batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest( $args );
			$result = $service->spreadsheets->batchUpdate($spreadsheet_id, $batchUpdateRequest);
			return true;
		}
		catch (Exception $e) {
			return false;
		}
	}

	public function remove_alternate_colors( $spreadsheet_id, $tab_id ) {
	
		try {
			$client = $this->getInstance();	
			
			if( ! $client ) {
				return false;
			}
			
			$service = new Google_Service_Sheets($client);
			$work_sheets = $service->spreadsheets->get($spreadsheet_id);
			
			$range = array( 'sheetId' => $tab_id );
			$args = array();
			
			$range_exist = false;
			
			$banded_range_id = 100;
			if( $tab_id != 0 ) {
				$generate_banded_range_id = substr($tab_id, 0, 4);
				$banded_range_id = $generate_banded_range_id;
			}
			
			foreach ($work_sheets as $sheet) {
				$properties = $sheet->getProperties();			
				if( $properties->sheetId == $tab_id ) {				
					$bandedRanges = $sheet->bandedRanges;				
					foreach( $bandedRanges as $bandedRange	 ) {					
						if( $bandedRange->bandedRangeId == $banded_range_id ) {
							$range_exist = true;
						}
					}
				}
			}
			
			if( $range_exist ) {
				$args = array( 
					'deleteBanding' => array(
						"bandedRangeId" => $banded_range_id,
					)
				);
				
				$banding_request = new Google_Service_Sheets_Request( $args );	
				$request = array( $banding_request );			
				
				$args = array( "requests" => $request );
				$batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest( $args );
				$result = $service->spreadsheets->batchUpdate($spreadsheet_id, $batchUpdateRequest);
				return true;
			}
		}
		catch (Exception $e) {
			return false;
		}
	}
	
	public function gsheet_get_google_account() {		
	
		try {
			$client = $this->getInstance();
			
			if( ! $client ) {
				return false;
			}
			
			$service = new Google_Service_Oauth2($client);
			$user = $service->userinfo->get();			
		}
		catch (Exception $e) {
			return false;
		}
		
		return $user;
	}
	
	public function gsheet_get_google_account_email() {		
		$google_account = $this->gsheet_get_google_account();	
		
		if( $google_account ) {
			return $google_account->email;
		}
		else {
			return "";
		}
	}
	
	public function gsheet_create_google_sheet($sheet_title = "") {
	
		$response = false;
		
		try {
			$client = $this->getInstance();
			
			if( ! $client ) {
				return false;
			}
			
			
			$title = $sheet_title ? $sheet_title : "GSheetConnector GravityForms";
			
			$properties = new Google_Service_Sheets_SpreadsheetProperties();
			$properties->setTitle($title);

			$spreadsheet = new Google_Service_Sheets_Spreadsheet();
			$spreadsheet->setProperties($properties);

			$sheet_service = new Google_Service_Sheets($client);		
			$create_spreadsheet = $sheet_service->spreadsheets->create( $spreadsheet );
			
			$spreadsheet = array(
				"spreadsheet_id" => $create_spreadsheet->spreadsheetId,
				"spreadsheet_name" => $title,
				"spreadsheet" => $create_spreadsheet,
				
			);
			$response = array( "result" => true, "spreadsheet" => $spreadsheet );
			
			do_action("gsheet_after_create_google_sheet", $response);
			
			$this->update_google_spreadsheets_option( $create_spreadsheet->spreadsheetId, $sheet_title );			
			// $this->sync_with_google_account();
		}
		catch (Exception $e) {
			$response = array( "result" => false, "error" => $e->getMessage() );
		}
		
		return $response;
	}


   /** 
   * GFGSC_googlesheet::gsheet_print_google_account_email
   * Get Google Account Email
   * @since 3.1 
   * @retun string $google_account
   **/
   public function gsheet_print_google_account_email() {

      try{
         $google_account = get_option("wpgs_email_account");
         if( false && $google_account ) {
            return $google_account;
         }
         else {  
            $google_sheet = new GSCWOO_googlesheet();
            $google_sheet->auth();            
            $email = $google_sheet->gsheet_get_google_account_email();
            update_option("wpgs_email_account", $email);
            return $email;
         }
      }catch(Exception $e){
         return false;
      }    
   }
	
}
