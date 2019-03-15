const bodyParser = require('body-parser')













var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

/* POST /forward */
router.get('/forward', async function(req, res, next) {
    const accessToken = await authHelper.getAccessToken(req.cookies, res);
    params.id = mail.getid();
    parms.debug = parms.getid;
  
    if (accessToken && id) {  
      // Initialize Graph client
      const client = graph.Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        }
      });
  
      try {

alert("ERROR");




        // const forward_params = {

            
        //   headers : { "content-type" : "application/json" ,  "authorization" : accessToken},
        //   body : JSON.stringify({
        //     "Comment": "REPORT",
        //     "ToRecipients": [
        //     {"EmailAddress": {
        //     "Address": "phishingreport@outlook.com"}}]
        //     });
        //   method : "POST",
        // }
        // const result = await client
        // .api('https://graph.microsoft.com/v1.0/me/messages/'+{id}+'/forward')
        // var data = JSON.stringify({
        // "Comment": "REPORT",
        // "ToRecipients": [
        // {"EmailAddress": {
        // "Address": "phishingreport@outlook.com"}}]
        // });
        // //.POST(postData);

      } catch (err) {
        parms.message = 'Error Reporting Message';
        parms.error = { status: `${err.code}: ${err.message}` };
        parms.debug = JSON.stringify(err.body, null, 2);
        res.render('error', parms);
      }
    } else {
      // Redirect to home
      res.redirect('/mail');
    }
    res.render('forward', parms);
  });

module.exports = router;





























// var express = require('express');
// var router = express.Router();
// var authHelper = require('../helpers/auth');
// var graph = require('@microsoft/microsoft-graph-client');


// function forward(id, accessToken){
// const id = id;
// const querystring = require('querystring');                                                                                                                                                                                                
// const https = require('https');

// var postData = JSON.stringify({
//   "Comment": "REPORT",
//   "ToRecipients": [
//   {"EmailAddress": {
//   "Address": "phishingreport@outlook.com"}}]
// });

// var options = {
//   hostname: 'graph.microsoft.com/v1.0',
//   path: '/me/messages/'+id+'/forward/',
//   method: 'POST',
//   headers: {
//        'Content-Type': 'application/JSON',
//        'Content-Length': postData.length,
//        'accessToken': accessToken
//      }
// };

// var req = https.request(options, (res) => {
//   console.log('statusCode:', res.statusCode);
//   console.log('headers:', res.headers);

//   res.on('data', (d) => {
//     process.stdout.write(d);
//   });
// });

// req.on('error', (e) => {
//   console.error(e);
// });

// req.write(postData);
// req.end();
// }







// POST /me/messages/+id+/forward/
// Host: graph.microsoft.com/v1.0
// Content-Type: application/JSON
// client_id= 





// var express = require('express');
// var router = express.Router();
// var authHelper = require('../helpers/auth');
// var graph = require('@microsoft/microsoft-graph-client');

// /*/POST//FORWARD FUNCTION*/
// function forward({id} , {accessToken}){
//   const headers = new headers()
//   headers.append('Content-Type', 'application/json');
//   const options = {
//     method = 'POST',
//     headers = accessToken + application/JSON,
//     data = {  
//   "Comment": "REPORT",
//   "ToRecipients": [
//     {
//       "EmailAddress": {
//         "Address": "phishingreport@outlook.com"
//       }
//     }
//   ]
//     },
//     body: JSON.stringify(data),
//   };
//   const id = {id};
//   const request = new Request ('https://graph.microsoft.com/v1.0/me/messages/' +{id} +'/forward',);

//   const resp = await fetch(req);
//   const status = await resp.status;

//   if (status === (201)) {
//     res.redirect('/mail/');
//     alert("Report Sent!");
//   }
//   else {
//     alert("Error Sending Report, Please Try Again!")
//   }
// }