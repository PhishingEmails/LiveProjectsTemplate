/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */



$(document).ready(() => {
    $('#run').click(run);
});
  
// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

async function submit() {
    /**
         * Insert your Outlook code here
         * 
         */

    POST https://outlook.office.com/api/v2.0/me/messages/AAMkAGE0Mz8DmAAA=/forward
    Content - Type: application / json

    {
        "Comment": "FYI",
            "ToRecipients": [
                {
                    "EmailAddress": {
                        "Address": "PhishingEmails.123@gmail.com"
                    }
                },
                {
                    "EmailAddress": {
                        "Address": "PhishingEmails.123@gmail.com"
                    }

                    // https://docs.microsoft.com/en-us/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#ForwardDirectly
                    //READ THIS
                }
            ]
    }


    

}