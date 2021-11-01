declare var fabric : any;
(function () {
    "use strict";
    
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            //messageBanner = new components.MessageBanner(element);
            //messageBanner.hideBanner();

            $('#get-data-from-selection').click(getDataFromSelection);
            $("#file").change(() => tryCatch(useInsertSlidesApi));
        });
    };
    
    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            }
        );
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    async function useInsertSlidesApi() {
        const myFile = <HTMLInputElement>document.getElementById("file");
        const reader = new FileReader();

        reader.onload = async (event) => {
            // strip off the metadata before the base64-encoded string
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            await PowerPoint.run(async function (context) {
                context.presentation.insertSlidesFromBase64(copyBase64);
                context.sync();
            });
        };

        // read in the file as a data URL so we can parse the base64-encoded string
        reader.readAsDataURL(myFile.files[0]);
    }

    /** Default helper for invoking an action and handling errors. */
    async function tryCatch(callback) {
        try {
            await callback();
        } catch (error) {
            // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
            console.error(error);
        }
    }

})();
