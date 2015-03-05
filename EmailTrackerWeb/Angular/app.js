

angular.module('emailTracker', []);


Office.initialize = function (reason) {
    $(document).ready(function () {
        console.log('Office Initialized');
        angular.bootstrap($('#container'), ['emailTracker']);
        console.log('Angular Initialized');
    });
};