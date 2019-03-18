var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');
var requestify = require('requestify');


/* GET /mail */
router.get('/', async function(req, res, next) {
    let parms = { title: 'Inbox', active: { inbox: true } };
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
        // Get the 25 newest messages from inbox
        const result = await client
        .api('/me/mailfolders/inbox/messages')
        .top(25)
        .select('subject,from,receivedDateTime,isRead,id')
        .orderby('receivedDateTime DESC')
        .get();
  
        parms.messages = result.value;
        res.render('mail', parms);
      } catch (err) {
        parms.message = 'Error retrieving messages';
        parms.error = { status: `${err.code}: ${err.message}` };
        parms.debug = JSON.stringify(err.body, null, 2);
        res.render('error', parms);
      }
  
    } else {
      // Redirect to home
      res.redirect('/');
    }
  });

  module.exports = router;




// function forward (id, accessToken){
//   requestify.request('https://graph.microsoft.com/v1.0/me/messages/'+{id}+'/forward', {
//     method: 'POST',
//     body: {
//       "Comment": "REPORT",
//        "ToRecipients": [
//            {"EmailAddress": {
//            "Address": "phishingreport@outlook.com"}}]
//            },
//     headers: {
//        'content-type' : "application/json" ,  "authorization" : 'accessToken'
//     },
//     dataType: 'json'        
// })
// .then(function(response) {
//     // get the response body
//     response.getBody();

//     // get the response headers
//     response.getHeaders();

//     // get specific response header
//     response.getHeader('Accept');

//     // get the code
//     response.getCode();

//     // Get the response raw body
//     response.body;
// });
// }





//   requestify.post('https://graph.microsoft.com/v1.0/me/messages/'+{id}+'/forward', {
//     hello: 'world'
// })
// .then(function(response) {
//     // Get the response body
//     response.getBody();
// });