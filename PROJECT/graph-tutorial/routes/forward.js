var express = require('express');
var router = express.Router();
var tokens = require('../tokens.js');
var graph = require('../graph.js');

/* GET /forward */
router.get('/',
  async function(req, res) {
    if (!req.isAuthenticated()) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      let params = {
        active: { forward: true }
      };

      // Get the access token
      var accessToken;
      try {
        accessToken = await tokens.getAccessToken(req);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err)
        });
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Get the events
          var forward = await graph.forwardEmail(accessToken);
          params.forward = forwardEmail.value;
        } catch (err) {
          req.flash('error_msg', {
            message: 'Could forward message',
            debug: JSON.stringify(err)
          });
        }
      }

      res.render('forward', params);
    }
  }
);

module.exports = router;