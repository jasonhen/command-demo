// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/// <reference path="../App.js" />



(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
          //  $('button').click(setLabel);
        });
    };

    function setLabel() {
        //  $('#statusLabel').html("Cheer sent!");
    //    loadMessageProps(Office.context.mailbox.item);
    }


    // Take an array of AttachmentDetails objects and
    // build a list of attachment names, separated by a line-break
    function buildAttachmentsString(attachments) {
        if (attachments && attachments.length > 0) {
            var returnString = "";

            for (var i = 0; i < attachments.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + attachments[i].name;
            }

            return returnString;
        }

        return "None";
    }

    // Format an EmailAddressDetails object as
    // GivenName Surname <emailaddress>
    function buildEmailAddressString(address) {
        return address.displayName;
    }

    // Take an array of EmailAddressDetails objects and
    // build a list of formatted strings, separated by a line-break
    function buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }

    // Load properties from a Message object
    function loadMessageProps(item) {
        $('#statusLabel').html("Message sent to: " + buildEmailAddressString(item.from));
    }

})();