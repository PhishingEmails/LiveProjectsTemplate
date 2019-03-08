function forward(message_id)
{
  POST https://graph.microsoft.com/v1.0/me/messages/{id}/forward
  Content-type: application/json
  Content-length: 166
  
  {
    "comment": "comment-value",
    "toRecipients": [
      {
        "emailAddress": {
          "name": "name-value",
          "address": "address-value"
        }
      }
    ]
  }

}







      // try {
      //   const result = await client
      //   .api('/me/messages/{message_id}/forward')
      //   .post();
      // }
    
