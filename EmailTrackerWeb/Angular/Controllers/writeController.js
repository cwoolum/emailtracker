angular.module('emailTracker').controller('writeController', ['$scope', function ($scope) {


    $scope.setBody = function () {
        var emailIdentifier = guid();
        var encodedData = window.btoa(JSON.stringify({
            "email_id": emailIdentifier

        })); // encode a string

        console.log(Office.context.mailbox.item.itemId);

        var imageTag = '<img src="https://api.keen.io/3.0/projects/54f7aa2b672e6c28c07fd8a5/events/email_opened?api_key=' +
        '7a5470f5e677d26d1d2a5e654a3ebdf76ae9874dcf45a63f8842f925153d99c6df36ce1b6795f955a1ba4af975ce475884727c98b1c7b9f1fa10ca07dac654e4b89b846ea97a7843aa3dc00dc3b2c7e818ccb4fa2ebe7de9397022f501ee15c4370be69f3baeb34c38738b55ff9ac4ac' +
        '&data=' + encodedData + '" class="keenImg" data-keen-id="' + emailIdentifier + '"/>'

        Office.cast.item.toItemCompose(Office.context.mailbox.item).body.setSelectedDataAsync(imageTag, {
            coercionType: Office.CoercionType.Html
        });
    };

    function guid() {
        function s4() {
            return Math.floor((1 + Math.random()) * 0x10000)
              .toString(16)
              .substring(1);
        }
        return s4() + s4() + '-' + s4() + '-' + s4() + '-' +
          s4() + '-' + s4() + s4() + s4();
    }


}]);