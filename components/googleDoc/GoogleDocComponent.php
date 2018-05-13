<?php
/**
 * Created by PhpStorm.
 * User: zvinger
 * Date: 13.05.18
 * Time: 13:17
 */

namespace Zvinger\GoogleDoc\components\googleDoc;

use Google_Client;
use Google_Service_Sheets;
use Google_Service_Sheets_BatchUpdateSpreadsheetRequest;

class GoogleDocComponent
{
    /**
     * Returns an authorized API client.
     * @return Google_Client the authorized client object
     * @throws \Google_Exception
     */
    public $clientSecretPath;

    /**
     * @return Google_Client
     * @throws \Google_Exception
     */
    public $credentialsPath;

    /**
     * @return Google_Client
     * @throws \Google_Exception
     */
    public function getClient()
    {
        $client = new Google_Client();
        $client->setApplicationName('Google Sheets API PHP Quickstart');
        $client->setScopes([Google_Service_Sheets::SPREADSHEETS]);
        $this->clientSecretPath = \Yii::getAlias("@runtime/google/client_secret.json");
        $client->setAuthConfig($this->clientSecretPath);
        $client->setAccessType('offline');
        if (file_exists($this->credentialsPath)) {
            $accessToken = json_decode(file_get_contents($this->credentialsPath), TRUE);
        } else {
            // Request authorization from the user.
            $authUrl = $client->createAuthUrl();
            printf("Open the following link in your browser:\n%s\n", $authUrl);
            print 'Enter verification code: ';
            $authCode = trim(fgets(STDIN));

            // Exchange authorization code for an access token.
            $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);

            // Store the credentials to disk.
            if (!file_exists(dirname($this->credentialsPath))) {
                mkdir(dirname($this->credentialsPath), 0700, TRUE);
            }
            file_put_contents($this->credentialsPath, json_encode($accessToken));
            printf("Credentials saved to %s\n", $this->credentialsPath);
        }
        $client->setAccessToken($accessToken);

        // Refresh the token if it's expired.
        if ($client->isAccessTokenExpired()) {
            $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            file_put_contents($this->credentialsPath, json_encode($client->getAccessToken()));
        }

        return $client;
    }

    private $activeSheet;

    /**
     * @param mixed $activeSheet
     * @return GoogleDocComponent
     */
    public function setActiveSheet($activeSheet)
    {
        $this->activeSheet = $activeSheet;

        return $this;
    }

    private $activeTab;

    public function updateRange($range, $rows)
    {
        $updates = new \Google_Service_Sheets_ValueRange();
        $updates->setValues($rows);

        return $this->getGoogleSheetService()->spreadsheets_values->update($this->activeSheet, $this->activeTab . $range, $updates, [
            'valueInputOption' => 'RAW',
        ]);
    }

    public function updateRangeByCoordinates($startColumn, $startRow, $endColumn, $endRow, $rows)
    {
        $stringRange = $this->num2alpha($startColumn) . ($startRow + 1) . ':' . $this->num2alpha($endColumn) . ($endRow + 1);

        return $this->updateRange($stringRange, $rows);
    }

    public function clearRange($range)
    {
        return $this->getGoogleSheetService()->spreadsheets_values->clear($this->activeSheet, $this->activeTab . $range, new \Google_Service_Sheets_ClearValuesRequest());
    }

    public function getRange($range)
    {
        return $this->getGoogleSheetService()->spreadsheets_values->get($this->activeSheet, $this->activeTab . $range);
    }

    public function getGoogleSheetService()
    {
        return new Google_Service_Sheets($this->getClient());
    }

    public function mergeRange($sheetId, $startColumn, $startRow, $endColumn, $endRow)
    {
        $mergeCells = new \Google_Service_Sheets_MergeCellsRequest();
        $rangeMerge = new \Google_Service_Sheets_GridRange();
        $rangeMerge->setSheetId($sheetId);
        $rangeMerge->setStartRowIndex($startRow);
        $rangeMerge->setStartColumnIndex($startColumn);
        $rangeMerge->setEndRowIndex($endRow);
        $rangeMerge->setEndColumnIndex($endColumn);
        $mergeCells->setRange($rangeMerge);
        $mergeCells->setMergeType('MERGE_ALL');
        $request = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest();
        $request->setRequests([
            [
                'mergeCells' => $mergeCells,
            ],
        ]);

        return $this->getGoogleSheetService()->spreadsheets->batchUpdate($this->activeSheet, $request);
    }

    public function freezeSheetInfo($tabId, $column, $row)
    {
        $google_Service_Sheets_BatchUpdateSpreadsheetRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest();
        $google_Service_Sheets_UpdateSheetPropertiesRequest = new \Google_Service_Sheets_UpdateSheetPropertiesRequest();
        $properties = new \Google_Service_Sheets_SheetProperties();
        $properties->setSheetId($tabId);
        $gridProperties = new \Google_Service_Sheets_GridProperties();
        $gridProperties->setFrozenColumnCount($column);
        $gridProperties->setFrozenRowCount($row);
        $properties->setGridProperties($gridProperties);
        $google_Service_Sheets_UpdateSheetPropertiesRequest->setProperties($properties);
        $google_Service_Sheets_UpdateSheetPropertiesRequest->setFields("gridProperties(frozenRowCount,frozenColumnCount)");
        $google_Service_Sheets_BatchUpdateSpreadsheetRequest->setRequests([
            [
                'updateSheetProperties' => $google_Service_Sheets_UpdateSheetPropertiesRequest,
            ],
        ]);
        $this->getGoogleSheetService()->spreadsheets->batchUpdate($this->activeSheet, $google_Service_Sheets_BatchUpdateSpreadsheetRequest);
    }

    /**
     * @param mixed $activeTab
     * @return GoogleDocComponent
     */
    public function setActiveTab($activeTab)
    {
        $this->activeTab = $activeTab . '!';

        return $this;
    }

    /**
     * @param $tabName
     * @return mixed
     * @throws \Exception
     */
    public function getSheetId($tabName, $create = FALSE)
    {
        $sheets = $this->getGoogleSheetService()->spreadsheets->get($this->activeSheet);
        /** @var \Google_Service_Sheets_Sheet $sheet */
        foreach ($sheets as $sheet) {
            $google_Service_Sheets_SheetProperties = $sheet->getProperties();
            $title = $google_Service_Sheets_SheetProperties->getTitle();
            if ($title == $tabName) {
                return $google_Service_Sheets_SheetProperties->getSheetId();
            }
        }
        if ($create) {
            return $this->createTab($tabName);
        }


        throw new \Exception("Google Sheet Tab " . $tabName . ' Form Sheet ID ' . $this->activeSheet . " not found");
    }

    public function createTab($tabName)
    {
        $sheet = new \Google_Service_Sheets_Sheet();
        $properties = new \Google_Service_Sheets_SheetProperties();
        $properties->setTitle($tabName);
        $sheet->setProperties($properties);
        $request = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest();
        $request->setRequests([
            [
                'addSheet' => [
                    'properties' => $properties,
                ],
            ],
        ]);
        $result = $this->getGoogleSheetService()->spreadsheets->batchUpdate($this->activeSheet, $request);

        return ($result->getReplies()[0]->getAddSheet()->getProperties()->sheetId);
    }

    public function num2alpha($n)
    {
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41) . $r;
        }

        return $r;
    }

    /**
     * @return mixed
     */
    public function getActiveSheet()
    {
        return $this->activeSheet;
    }
}