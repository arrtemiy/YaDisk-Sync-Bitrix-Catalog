<?php
if (substr(php_sapi_name(), 0, 3) !== 'cli') {
    die();
}
define('NO_KEEP_STATISTIC', true);
define('NO_AGENT_CHECK', true);

if (empty($_SERVER['DOCUMENT_ROOT'])) {
    $_SERVER['DOCUMENT_ROOT'] = '/home/bitrix/www';
}

require($_SERVER['DOCUMENT_ROOT'] . '/bitrix/modules/main/include/prolog_before.php');
require_once $_SERVER['DOCUMENT_ROOT'] . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

$accessToken = 'y0_AgAEA7qjsQ4FAADLWwAAAAELS7AQAACtODO3Zy5GZaEJ08cj28W9oDXhNg';
$savePath = $_SERVER['DOCUMENT_ROOT'] . '/bitrix/php_interface/include/quantity_upd/Pedrollo.xlsx';

// Получаем прямую ссылку на файл
$fileUrl = 'https://cloud-api.yandex.net/v1/disk/resources/download?path=%2F%D0%9E%D1%81%D1%82%D0%B0%D1%82%D0%BA%D0%B8%2FPedrollo.xlsx';
$headers = [
    "Authorization: OAuth $accessToken"
];

$response = getCurlResponse($fileUrl, $headers);

if (!$response['success']) {
    addMsg2Log("Ошибка получения УРЛ: " . $response['error']);
} else {
    $responseData = json_decode($response['data'], true);
    if (isset($responseData['error'])) {
        addMsg2Log("API Error: " . $responseData['error']);
    } else {
        $downloadUrl = $responseData['href'];

        // Скачиваем файл по прямой ссылке с увеличенным таймаутом
        $fileResponse = getCurlResponse($downloadUrl);
        if (!$fileResponse['success']) {
            addMsg2Log("Ошибка загрузки: " . $fileResponse['error']);
        } else {
            file_put_contents($savePath, $fileResponse['data']);
        }
    }
}

// Чтение и обработка файла
$spreadsheet = IOFactory::load($savePath);
$sheet = $spreadsheet->getActiveSheet();
$highestRow = $sheet->getHighestRow();
$highestColumn = $sheet->getHighestColumn();

$articles = [];
for ($row = 2; $row <= $highestRow; $row++) {
    $article = trim($sheet->getCell('A' . $row)->getValue());
    $quantity = $sheet->getCell('J' . $row)->getValue();
    if (empty($quantity)) {
        $quantity = 0;
    }
    if (strlen($article) >= 3) {
        $articles[$article] = $quantity;
    }
}

// Обновляем остатки
CModule::IncludeModule('iblock');

$iblockIds = [42, 61];

foreach ($iblockIds as $iblockId) {
    // Получаем элементы инфоблока
    $arSelect = ["ID", "IBLOCK_ID", "PROPERTY_CML2_ARTICLE"];
    $arFilter = ["IBLOCK_ID" => $iblockId];
    $res = CIBlockElement::GetList([], $arFilter, false, false, $arSelect);

    while ($ob = $res->GetNextElement()) {
        $arFields = $ob->GetFields();
        $article = $arFields['PROPERTY_CML2_ARTICLE_VALUE'];

        // Ищем артикул в массиве
        if (isset($articles[$article]) && !empty($article)) {
            $quantityAvailable = $articles[$article];

            // Получаем ID товара по артикулу
            $productId = CCatalogProduct::GetByID($arFields['ID']);
            if ($productId) {
                // Обновляем "Доступное количество" (CAT_BASE_QUANTITY)
                $updateProductFields = [
                    'QUANTITY' => $quantityAvailable
                ];

                $productUpdateResult = CCatalogProduct::Update($productId['ID'], $updateProductFields);

                // Обновляем "Кол-во товара на складе"
                $storeFields = [
                    'PRODUCT_ID' => $productId['ID'],
                    'STORE_ID' => 1,
                    'AMOUNT' => $quantityAvailable
                ];

                $storeUpdateResult = CCatalogStoreProduct::UpdateFromForm($storeFields);
                if ($storeUpdateResult) {
                    addMsg2Log($productId['ID'] . " Количество на складе обновлено успешно.");
                } else {
                    addMsg2Log("Ошибка при обновлении количества на складе для товара ID: " . $productId['ID']);
                }
            } else {
                addMsg2Log("Товар не найден или неверный ID.");
            }
        }
    }
}

function getCurlResponse($url, $headers = [], $timeout = 120) {
    $ch = curl_init($url);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
    curl_setopt($ch, CURLOPT_TIMEOUT, $timeout);
    curl_setopt($ch, CURLOPT_CONNECTTIMEOUT, 60);
    curl_setopt($ch, CURLOPT_IPRESOLVE, CURL_IPRESOLVE_V4);

    if (!empty($headers)) {
        curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
    }

    $response = curl_exec($ch);
    if (curl_errno($ch)) {
        $error = curl_error($ch);
        curl_close($ch);
        return ['success' => false, 'error' => $error];
    }

    curl_close($ch);
    return ['success' => true, 'data' => $response];
}

require($_SERVER['DOCUMENT_ROOT'] . '/bitrix/modules/main/include/epilog_after.php');