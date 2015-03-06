angular.module('emailTracker').controller('readController', ['$q', '$scope', function ($q, $scope) {


    function hasOpened(email_id, callback) {
        var keen_client = new Keen({
            projectId: "54f7aa2b672e6c28c07fd8a5", // String (required always)
            readKey: "3216e90488814a5a97f9b8898248e26a50d7919f4472ffccc479b68230b1e6b4eadf8564a760f3e2da896d7afb27aed85bbc1c98348d78ca6007192be55db69708f18b41e7bef78a9f4510a4fba5e38e969181348841dd2a180da00f3c8323809be8f2d5b99098276fad7fa6cbb4b652", // String (required for querying data)
            protocol: "auto", // String (optional: https | http | auto)
            host: "api.keen.io/3.0", // String (optional)
            requestType: "jsonp" // String (optional: jsonp, xhr, beacon)
        });

        var count_query = new Keen.Query("count", {
            eventCollection: "email_opened",
            groupBy: "email_id",
            filters: [
              {
                  "property_name": "email_id",
                  "operator": "eq",
                  "property_value": email_id
              }]
        });

        // Send query
        keen_client.run(count_query, function (err, res) {
            if (err) {
                console.log(err);
            }
            else {
                callback(res.result[0].result > 1);
            }
        });
    };

    function getKeenIdFromEmail() {
        var deferred = $q.defer();

        try {
            var currentEmail = Office.cast.item.toItemRead(Office.context.mailbox.item);

            deferred.resolve(currentEmail.getRegExMatches().KeenId);
        } catch (error) {
            deferred.reject(error);
        }

        return deferred.promise;
    }

    getKeenIdFromEmail().then(function(keenId) {
        hasOpened(keenId, function (result) {
            $scope.result = result ? "opened" : "not opened";
            console.log(result);
        });
    });




}]);