
function forward(){
  const result = await client
  const api = { '/me/messages/':id+'/forward'}
  .api(api)
  .post();
  res.send(request)
}



function qforward() {


var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

/* POST /forward */
router.post('/submit', function(req, res, next){
  var id = req.params.id
  res.redirect('/submit/' + id)


  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

      if (accessToken && userName) {
      parms.user = userName;
  
      // Initialize Graph client
      const client = graph.Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        }
      });
  
      try {
        const result = await client
        const api = { '/me/messages/':id+'/forward'}
        .api(api)
        .post();
        res.send(request)

  
        parms.messages = result.value;
        res.render('mail', parms);
      } catch (err) {
        parms.message = 'Error Sending Report';
        parms.error = { status: `${err.code}: ${err.message}` };
        parms.debug = JSON.stringify(err.body, null, 2);
        res.render('error', parms);
      }
  
    } else {
      // Redirect to home
      res.redirect('/mail/');
    }
  });

module.exports = router

}






// var express = require('express');
// var router = express.Router();
// var authHelper = require('../helpers/auth');
// var graph = require('@microsoft/microsoft-graph-client');

// /* POST /forward */
// router.post('/', async function(req, res, next) {
//     let parms = { title: 'Report', active: { Report: true } };
//     const accessToken = await authHelper.getAccessToken(req.cookies, res);
//     const userName = req.cookies.graph_user_name;
  
//     if (accessToken && userName) {
//       parms.user = userName;
  
//       // Initialize Graph client
//       const client = graph.Client.init({
//         authProvider: (done) => {
//           done(null, accessToken);
//         }
//       });
  
//       try {
//         const result = await client
//         cont id = await {{this.id}}
//         .api('/me/messages/{message_id}/forward')
//         .post();

//         res.send(request)
  
//         parms.messages = result.value;
//         res.render('mail', parms);
//       } catch (err) {
//         parms.message = 'Error Sending Report';
//         parms.error = { status: `${err.code}: ${err.message}` };
//         parms.debug = JSON.stringify(err.body, null, 2);
//         res.render('error', parms);
//       }
  
//     } else {
//       // Redirect to home
//       res.redirect('/mail/');
//     }
//   });

// module.exports = router;





























// private static $outlookApiUrl = "https://outlook.office.com/api/v2.0";

// public static function makeApiCall($access_token, $user_email, $method, $url, $payload = NULL) {
//     // Generate the list of headers to always send.

//     $headers = array(
//         "User-Agent: php-tutorial/1.0", // Sending a User-Agent header is a best practice.
//         "Authorization: Bearer " . $access_token, // Always need our auth token!
//         "Accept: application/json", // Always accept JSON response.
//         "client-request-id: " . self::makeGuid(), // Stamp each new request with a new GUID.
//         "return-client-request-id: true", // Tell the server to include our request-id GUID in the response.
//         "X-AnchorMailbox: " . $user_email         // Provider user's email to optimize routing of API call
//     );

//     $curl = curl_init($url);

//     switch (strtoupper($method)) {
//         case "GET":
//             // Nothing to do, GET is the default and needs no
//             // extra headers.
//             error_log("Doing GET");
//             break;
//         case "POST":
//             error_log("Doing POST");
//             // Add a Content-Type header (IMPORTANT!)
//             $headers[] = "Content-Type: application/json";
//             curl_setopt($curl, CURLOPT_POST, true);
//             curl_setopt($curl, CURLOPT_POSTFIELDS, $payload);
//             break;
//         case "PATCH":
//             error_log("Doing PATCH");
//             // Add a Content-Type header (IMPORTANT!)
//             $headers[] = "Content-Type: application/json";
//             curl_setopt($curl, CURLOPT_CUSTOMREQUEST, "PATCH");
//             curl_setopt($curl, CURLOPT_POSTFIELDS, $payload);
//             break;
//         case "DELETE":
//             error_log("Doing DELETE");
//             curl_setopt($curl, CURLOPT_CUSTOMREQUEST, "DELETE");
//             break;
//         default:
//             error_log("INVALID METHOD: " . $method);
//             exit;
//     }

//     curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
//     curl_setopt($curl, CURLOPT_HTTPHEADER, $headers);
//     $response = curl_exec($curl);
//     error_log("curl_exec done.");
//     //echo "test";print_r($response);
//     $httpCode = curl_getinfo($curl, CURLINFO_HTTP_CODE);
//     //echo "test".$httpCode;
//     error_log("Request returned status " . $httpCode);

//     if ($httpCode >= 400) {
//         return array('errorNumber' => $httpCode,
//             'error' => 'Request returned HTTP error ' . $httpCode);
//     }

//     $curl_errno = curl_errno($curl);
//     $curl_err = curl_error($curl);

//     if ($curl_errno) {
//         $msg = $curl_errno . ": " . $curl_err;
//         error_log("CURL returned an error: " . $msg);
//         curl_close($curl);
//         return array('errorNumber' => $curl_errno,
//             'error' => $msg);
//     } else {
//         error_log("Response: " . $response);
//         curl_close($curl);
//         return json_decode($response, true);
//     }
// }



