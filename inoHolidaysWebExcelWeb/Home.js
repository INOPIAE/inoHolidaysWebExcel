(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // Die Initialisierungsfunktion muss bei jedem Laden einer neuen Seite ausgeführt werden.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Benachrichtigungsmechanismus initialisieren und ausblenden
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // Wenn nicht Excel 2016 verwendet wird, Fallbacklogik verwenden.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#eyear-text").text("Year:");
                $("#template-description").text("Add Easter date to current cell.");
                $('#addEaster-text').text("Add Easter date");
                $('#addEaster-desc').text("Add Easter date to current cell.");

                $('#addEaster-button').click(addEaster);
                return;
            }

            $("#template-description").text("Add Easter date to current cell.");
            $('#addEaster-text').text("Add Easter date");
            $('#addEaster-desc').text("Add Easter date to current cell.");

            // Fügt einen Klickereignishandler für die Hervorhebungsschaltfläche hinzu.
            $('#addEaster-button').click(addEaster);
        });
    };

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('Der ausgewählte Text lautet:', '"' + result.value + '"');
                } else {
                    showNotification('Fehler', result.error.message);
                }
            });
    }


    function dateFromEaster(Year, DaysDiff) { 
        if ((Year == "") || (Year == null)) { Year = new Date().getFullYear() };
        
        if ((Year < 1970) || (2099 < Year)) { return "Year must be between 1970 and 2099"; };

        if ((DaysDiff == "") || (DaysDiff == null)) { DaysDiff = 0; };

        var a = Year % 19;
        var d = (19 * a + 24) % 30;
        var eday = d + (2 * (Year % 4) + 4 * (Year % 7) + 6 * d + 5) % 7;
        if ((eday == 35) || ((eday == 34) && (d == 28) && (a > 10))) { eday -= 7; };

        var EasterDate = new Date(Year, 2, 22);
       
        // 86400000 = 24 h * 60 min * 60 s * 1000 ms
        // 60000  = 60 s * 1000 ms
        EasterDate.setTime(EasterDate.getTime() + 86400000 * DaysDiff + 86400000 * eday - EasterDate.getTimezoneOffset() * 60000);

        EasterDate = EasterDate.toISOString();
        EasterDate = EasterDate.substring(0, 10);

        return EasterDate;
    }

    function addEaster() {
        Excel.run(function (ctx) {

            var year = document.getElementById("eyear").value;

            var range = ctx.workbook.getActiveCell()

            if (isNaN(year) == true) {
                range.values = "Year must be a number between 1970 and 2099";
            } else if (parseInt(year) < 1970 || parseInt(year) > 2099 ){
                range.values = "Year must be a number between 1970 and 2099";
            } else {
                var edate = moment(dateFromEaster(year, 0));

                range.values = edate.format('L');
            }
            return ctx.sync();
        })
            .catch(errorHandler);
    }

    // Eine Hilfsfunktion zur Behandlung von Fehlern.
    function errorHandler(error) {
        // Stellen Sie immer sicher, dass kumulierte Fehler abgefangen werden, die bei der Ausführung von "Excel.run" auftreten.
        showNotification("Fehler", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Eine Hilfsfunktion zum Anzeigen von Benachrichtigungen.
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
