/// <reference path="../App.js" />

(function () {
    'use strict';


    // The initialize function must be run each time a new page is loaded
    //Office.initialize = function (reason) {
    //    $(document).ready(function () {
    //        app.initialize();

    //        $('#set-subject').click(setSubject);
    //        $('#get-subject').click(getSubject);
    //        $('#set-body').click(setBody);
    //        $('#add-to-recipients').click(addToRecipients);
    //    });
    //};

    function setBody() {
        var emailIdentifier=guid();
        var encodedData = window.btoa(JSON.stringify({
            "email_id": emailIdentifier
           
        })); // encode a string

        console.log(Office.context.mailbox.item.itemId);

        var imageTag = '<img src="https://api.keen.io/3.0/projects/54f7aa2b672e6c28c07fd8a5/events/email_opened?api_key=' +
        '7a5470f5e677d26d1d2a5e654a3ebdf76ae9874dcf45a63f8842f925153d99c6df36ce1b6795f955a1ba4af975ce475884727c98b1c7b9f1fa10ca07dac654e4b89b846ea97a7843aa3dc00dc3b2c7e818ccb4fa2ebe7de9397022f501ee15c4370be69f3baeb34c38738b55ff9ac4ac' +
        '&data=' + encodedData + '" data-keen-id="'+emailIdentifier+'"/>'

        Office.cast.item.toItemCompose(Office.context.mailbox.item).body.setSelectedDataAsync(imageTag, {
            coercionType: Office.CoercionType.Html
        });
    }

    function guid() {
        function s4() {
            return Math.floor((1 + Math.random()) * 0x10000)
              .toString(16)
              .substring(1);
        }
        return s4() + s4() + '-' + s4() + '-' + s4() + '-' +
          s4() + '-' + s4() + s4() + s4();
    }

    function setSubject() {
        Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync("Hello world!");
    }

    function getSubject() {
        Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.getAsync(function (result) {
            app.showNotification('The current subject is', result.value)
        });
    }

    function addToRecipients() {
        var item = Office.context.mailbox.item;
        var addressToAdd = {
            displayName: Office.context.mailbox.userProfile.displayName,
            emailAddress: Office.context.mailbox.userProfile.emailAddress
        };

        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
        }
    }

})();