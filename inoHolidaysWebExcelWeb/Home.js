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
                $('#addHoliday-text').text("Add Holiday");
                $('#addHoliday-desc').text("Add public holidays starting in current cell.");

                $('#addEaster-button').click(addEaster);
                $('#addHoliday-button').click(addHolidays1);
                return;
            }

            $("#template-description").text("Add Easter date to current cell.");
            $('#addEaster-text').text("Add Easter date");
            $('#addEaster-desc').text("Add Easter date to current cell.");
            $('#addHoliday-text').text("Add Holiday");
            $('#addHoliday-desc').text("Add public holidays starting in current cell.");
            // Fügt einen Klickereignishandler für die Hervorhebungsschaltfläche hinzu.
            $('#addEaster-button').click(addEaster);
            $('#addHoliday-button').click(addHolidays);
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

    async function addHolidays() {
        await Excel.run(async (ctx) => {

            let activeCell = ctx.workbook.getActiveCell();
            activeCell.load("address");

            await ctx.sync();

            var celladdress = activeCell.address.split('!');

            var rangeAddress = celladdress[1];
            var sheetName = celladdress[0];

            let tableName = 'GermanHolidays';
            let tc = '';
            let table = '';

            let tableCount = ctx.workbook.tables.getCount();
            await ctx.sync();

            let tcount = tableCount.value;

            for (let i = 0; i < tcount; i++) {
                table = ctx.workbook.tables.getItemAt(i);
                table.load('name');
                await ctx.sync();
                if (table.name === tableName) {
                    tc++;
                    tableName = 'GermanHolidays' + tc;
                    i = -1;
                }
            }

            var year = document.getElementById("eyear").value;

            if (isNaN(year) == true) {
                range.values = "Year must be a number between 1970 and 2099";
            } else if (parseInt(year) < 1970 || parseInt(year) > 2099) {
                range.values = "Year must be a number between 1970 and 2099";
            } else {

                if ((year == "") || (year == null)) { year = new Date().getFullYear() };

                const holidays = [
                    ["01.01.2021", "Neujahr", "All"],
                    ["06.01.2021", "Drei Hl.Könige", "BY, ST, BW"],
                    ["08.03.2021", "Int.Frauentag", "BE, MV"],
                    ["Ostern - 2", "Karfreitag", "All"],
                    ["Ostern + 1", "Ostermontag", "All"],
                    ["01.05.2021", "Tag der Arbeit", "All"],
                    ["Ostern + 39", "Christi Himmelfahrt", "All"],
                    ["Ostern + 50", "Pfingstmontag", "All"],
                    ["Ostern + 60", "Fronleichnam", "BW, BY, HE, NW, RP, SL"],
                    ["15.08.2021", "Maria Himmelfahrt", "BY, SL"],
                    ["20.09.2021", "Weltkindertag", "TH"],
                    ["03.10.2021", "Tag der dt.Einheit", "All"],
                    ["31.10.2021", "Reformationstag", "SH, NI, HB, HH, BB, ST, SN, TH, MV"],
                    ["01.11.2021", "Allerheiligen", "BW, BY, NW, RP, SL"],
                    ["Advent - 32", "Buß - und Bettag", "SN"],
                    ["25.12.2021", "1. Weihnachtstag", "All"],
                    ["26.12.2021", "2. Weihnachtstag", "All"]
                ];

                let sheet = ctx.workbook.worksheets.getItem(sheetName);
                
                let holidayTable = sheet.tables.add(rangeAddress +":" + moveCell(rangeAddress, 0, 2), true /*hasHeaders*/);
                holidayTable.name = tableName;

                holidayTable.getHeaderRowRange().values = [["Date", "Name", "Regions/State"]];

                for (const holiday of holidays) {
                    holidayTable.rows.add(null /*add rows to the end of the table*/,
                        [[getHolidayDate(holiday[0], year), holiday[1], holiday[2] ]]);
                }

                if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                    sheet.getUsedRange().format.autofitColumns();
                    sheet.getUsedRange().format.autofitRows();
                }

            }
            return ctx.sync();
        })
            .catch(errorHandler);
    }

    function getHolidayDate(date, year) {
        const holiday = date.split(' ');
        if (holiday[0] === "Ostern") {
            var direction = 1;
            if (holiday[1] === "-") {
                direction = -1;
            }
            var edate = moment(dateFromEaster(year, direction * holiday[2]));
            return edate.format('L');
        } else if (holiday[0] === "Advent") {
            var direction = 1;
            if (holiday[1] === "-") {
                direction = -1;
            }
            const momentDate = lastAdvent(year);
            momentDate.add(direction * holiday[2], 'd');
            return momentDate.format('L');
        } else {
            const hdate = date.split('.');
            const hdateConv = year + '-' + hdate[1].padStart(2, "0") + '-' + hdate[0].padStart(2, "0");
            const momentDate = moment(hdateConv + 'T00:00:00.000+00:00')
            return momentDate.format('L');
        }
    }

    function lastAdvent(year) {
        const momentDate = moment(year + '-12-24T00:00:00.000+00:00')
        const wkday = momentDate.day();
        if (wkday != 0) {
            momentDate.add(-wkday, 'd');
        }
        return momentDate;
    }

    function moveCell(cell, rows, columns) {
        if ((rows == "") || (rows == null)) { rows = 0 };
        if ((columns == "") || (columns == null)) { columns = 0 };
        let cellparts = cell.match(/[a-zA-Z]+|[0-9]+/g);

        let char = cellparts[0].split('');
        let i = 0;
        let cv = 0;

        for (let i = 0 ; i < char.length ; i++) {
            cv += (char[char.length - 1 - i].charCodeAt() - 64) * (26 ** i)
        }

        cv += columns
        let result = '';
        i = char.length;
        do {
            i--;
            const quotient = Math.floor(cv / (26 ** i));   // 1
            cv = cv % (26 ** i); 
            result += String.fromCharCode(quotient + 64);
        } while (i > 0);

        result += (+cellparts[1] + rows)

        return result;
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
