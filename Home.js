
var app = angular.module('forwordApp', ['ngMaterial', 'ngRoute', 'ui.router',]);

app.controller('forwordAppCtrl', function ($scope, $mdToast, $log, $mdDialog) {


    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {

            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };
    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };
    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };


    function isTokenExpired(token) {
        const base64Url = token.split(".")[1];
        const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
        const jsonPayload = decodeURIComponent(
            atob(base64)
                .split("")
                .map(function (c) {
                    return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
                })
                .join("")
        );
        const { exp } = JSON.parse(jsonPayload);
        var expnew = exp * 1000;

        var ee = new Date(Date.now());
        var ef = new Date(expnew);

        if (new Date(Date.now()) > new Date(expnew)) {
            expired = true;
        }
        else {
            expired = false;
        }

        //console.log(expired);
        return expired
    };



    ProgressLinearActive();

    $scope.LoginDiv = false;
    $scope.MainDiv = true;

    var accessToken = window.localStorage.getItem("accessToken");
    if (accessToken != null && accessToken != undefined && accessToken != "") {
        $scope.LoginDiv = true;
        $scope.MainDiv = false;
        ProgressLinearInActive();
    }else {
        ProgressLinearInActive();
    };


    //if (accessToken) {
    //    var check_Expiration = isTokenExpired(accessToken);
    //    console.log(check_Expiration);
    //     if (check_Expiration) {
    //        $scope.SignP_Out();
    //    }
    //};

   
       

   // ShowLoader();

    Office.onReady(function () {


     //   console.log(accessToken)
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function (asyncResult) {

            mailbody_with_Formate = asyncResult.value;
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                // TODO: Handle error
            }
            else {

            }
        })



          


        var emailId = Office.context.mailbox.item.itemId;

        let dialog;
        $scope.login = function () {
            let a = "https://login.microsoftonline.com/bb7a26e1-3781-4e86-a15c-7e2b0f934951/oauth2/v2.0/authorize?client_id=3f623ba6-ead5-4cd3-9f64-d01ea250f581&response_type=token&redirect_uri=https://aamirhusnain.github.io/ForwardEmail/Redirect.html&scope=user.read%20mail.readwrite%20mail.send&response_mode=fragment&state=12345&nonce=678910";
            // let a = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=3f623ba6-ead5-4cd3-9f64-d01ea250f581&response_type=token&redirect_uri=https://localhost:44397/Template/Redirect.html&scope=user.read%20mail.readwrite%20mail.send&response_mode=fragment&state=12345&nonce=678910";
            //let a = "https://localhost:44397/Template/Redirect.html";

            // mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
            // Declare dialog as global for use in later functions.
            Office.context.ui.displayDialogAsync(a, { height: 50, width: 40 },
                function (asyncResult) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
            );
        }

        function processMessage(arg) {
            dialog.close();
            let messageFromDialog = JSON.parse(arg.message);
            window.localStorage.setItem("accessToken", messageFromDialog)
            window.location.reload();
        };
        $scope.SignP_Out = function () {
            window.localStorage.clear();
            window.location.reload();
        };


        $scope.forword_Mail = function () {

            ProgressLinearActive();

            var emailData = {
                'comment': "<table  border=1> <th>Customer</th> <th>Project</th>  <th>Country</th></tr >   <tr> <td >" + $scope.Customer + "</td>  <td>" + $scope.Project + "</td>  <td>" + $scope.Country + "</td></tr></table>",
                'toRecipients': [
                    {
                        "emailAddress": {
                            "address": $scope.toRecipient,

                        }
                    }
                ]
            };


            var settings = {
                "url": 'https://graph.microsoft.com/v1.0/me/messages/' + emailId + ' /forward',
                "method": "POST",
                "timeout": 0,
                "headers": {
                    "Authorization": "Bearer " + accessToken,
                    "Content-Type": "application/json"
                },
                "data": JSON.stringify(emailData),
            };

            $.ajax(settings).done(function (response) {
              //  console.log(response);

                $scope.Customer = undefined;
                $scope.toRecipient = undefined;
                $scope.Country = undefined;
                $scope.Project = undefined;

                ProgressLinearInActive();
                loadToast("Mail Forwarded");

            }).fail(function (error) {
                console.log(error);


                ProgressLinearInActive();
                loadToast("Request Failed");

            });
        };






        


    });
}).config(function ($mdThemingProvider) {
    $mdThemingProvider.theme('dark-grey').backgroundPalette('grey').dark();
    $mdThemingProvider.theme('dark-orange').backgroundPalette('orange').dark();
    $mdThemingProvider.theme('dark-purple').backgroundPalette('deep-purple').dark();
    $mdThemingProvider.theme('dark-blue').backgroundPalette('blue').dark();
});
