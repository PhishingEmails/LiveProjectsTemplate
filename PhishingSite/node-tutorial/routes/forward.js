var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');


function forward({id} , {accessToken}){
  const headers = new headers()
  headers.append('Content-Type', 'application/json');
  const options = {
    method = 'POST',
    headers = accessToken + application/JSON,
    data = {  
  "Comment": "REPORT",
  "ToRecipients": [
    {
      "EmailAddress": {
        "Address": "phishingreport@outlook.com"
      }
    }
  ]
    },
    //body: JSON.stringify(data),
  };
  const id = {id};
  const request = new Request ('https://graph.microsoft.com/v1.0/me/messages/' +{id} +'/forward',);

  const resp = await fetch(req);
  const status = await resp.status;

  if (status === (201)) {
    res.redirect('/mail/');
    alert("Report Sent");
  }
}