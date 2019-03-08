private static $outlookApiUrl = "https://outlook.office.com/api/v2.0";

public static function makeApiCall($access_token, $user_email, $method, $url, $payload = NULL) {
    // Generate the list of headers to always send.

    $headers = array(
        "User-Agent: php-tutorial/1.0", // Sending a User-Agent header is a best practice.
        "Authorization: Bearer " . $access_token, // Always need our auth token!
        "Accept: application/json", // Always accept JSON response.
        "client-request-id: " . self::makeGuid(), // Stamp each new request with a new GUID.
        "return-client-request-id: true", // Tell the server to include our request-id GUID in the response.
        "X-AnchorMailbox: " . $user_email         // Provider user's email to optimize routing of API call
    );

    $curl = curl_init($url);

    switch (strtoupper($method)) {
        case "GET":
            // Nothing to do, GET is the default and needs no
            // extra headers.
            error_log("Doing GET");
            break;
        case "POST":
            error_log("Doing POST");
            // Add a Content-Type header (IMPORTANT!)
            $headers[] = "Content-Type: application/json";
            curl_setopt($curl, CURLOPT_POST, true);
            curl_setopt($curl, CURLOPT_POSTFIELDS, $payload);
            break;
        case "PATCH":
            error_log("Doing PATCH");
            // Add a Content-Type header (IMPORTANT!)
            $headers[] = "Content-Type: application/json";
            curl_setopt($curl, CURLOPT_CUSTOMREQUEST, "PATCH");
            curl_setopt($curl, CURLOPT_POSTFIELDS, $payload);
            break;
        case "DELETE":
            error_log("Doing DELETE");
            curl_setopt($curl, CURLOPT_CUSTOMREQUEST, "DELETE");
            break;
        default:
            error_log("INVALID METHOD: " . $method);
            exit;
    }

    curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($curl, CURLOPT_HTTPHEADER, $headers);
    $response = curl_exec($curl);
    error_log("curl_exec done.");
    //echo "test";print_r($response);
    $httpCode = curl_getinfo($curl, CURLINFO_HTTP_CODE);
    //echo "test".$httpCode;
    error_log("Request returned status " . $httpCode);

    if ($httpCode >= 400) {
        return array('errorNumber' => $httpCode,
            'error' => 'Request returned HTTP error ' . $httpCode);
    }

    $curl_errno = curl_errno($curl);
    $curl_err = curl_error($curl);

    if ($curl_errno) {
        $msg = $curl_errno . ": " . $curl_err;
        error_log("CURL returned an error: " . $msg);
        curl_close($curl);
        return array('errorNumber' => $curl_errno,
            'error' => $msg);
    } else {
        error_log("Response: " . $response);
        curl_close($curl);
        return json_decode($response, true);
    }
}