var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

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

function forward(){
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
      try{
        var data = 
      {
      "Comment": "Report",
      "ToRecipients": [{
      "EmailAddress": {
      "Address": "phishingreport@outlook.com"}}]
      }
      
      const api = { '/me/messages/':id+'/forward'}
      .api(api)
      .post(data);
      res.send(request)
      redirect('/home/')
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
}