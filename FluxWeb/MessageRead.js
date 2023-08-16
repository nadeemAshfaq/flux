


var app = angular.module('myApp', ['ngMaterial', 'ngRoute']);



app.controller('myAppCtrl', function ($scope, $mdToast, $mdDialog) {

   
    Office.onReady(function () {

        $scope.sendApiRequest = function (event) {

            var mailItem = Office.context.mailbox.item;

            var attachment = mailItem.attachments;
            var senderMail = mailItem.from.emailAddress;
            var cc = Office.context.mailbox.item.cc;


            var subject = mailItem.subject;
            var sender = mailItem.to[0];
            var timeCreated = mailItem.dateTimeCreated;
           /* var timereceived = mailItem.dateTimeModified;*/





           


            ////////////////////////////bodytext
            mailItem.body.getAsync(Office.CoercionType.Text, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var bodyText = result.value;
                    // Process the body text here
                    console.log("Mail Body Text:", bodyText);
                } else {
                    // Handle error
                    console.error("Error retrieving mail body:", result.error.message);
                }
            });





            // Create an object to hold the data
            var dataObject = {
                
                attachment: attachment,
                senderMail: senderMail,
                cc: cc,
                subject: subject,
                sender: sender,
              
                timeCreated: timeCreated,
                //timeReceived: timeReceived,
                bodyText: null // Placeholder for the body text
            };

            // Retrieve the body text
            mailItem.body.getAsync(Office.CoercionType.Text, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    // Set the retrieved body text in the dataObject
                    dataObject.bodyText = result.value;

                    // Convert the data object to a JSON string
                    var dataString = JSON.stringify(dataObject);



                    // Send the dataString to the Pipedream URL using fetch
                    fetch($scope.apiUrl, {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json",
                            "Authorization": "Bearer " + $scope.apiKey
                        },
                        body: dataString
                    })
                        .then(response => response.json())
                        .then(data => {
                            console.log(' Response:', data);
                            alert("data post successfully")
                            
                            $scope.hideLoader();
                        })
                        .catch(error => {
                            console.error('Error:', error);
                            alert(error);
                        });

                } else {
                    // Handle error retrieving body text
                    console.error("Error retrieving mail body:", result.error.message);
                }

            })




        }




    })



})