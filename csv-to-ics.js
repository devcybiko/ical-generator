const glstools = require("glstools");
const gfiles = glstools.files;
const gprocs = glstools.procs;

const KEY_WORDS = {
    HEADER: 'BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//ICS Generator//Tim Rose//EN\nCALSCALE:GREGORIAN\n',
    TITLE: 'X-WR-CALNAME;VALUE=TEXT:',
    EVENT: ['BEGIN:VEVENT', 'END:VEVENT'],
    FOOTER: 'END:VCALENDAR',
    BUSY_STATUS: ['FREE', 'WORKINGELSEWHERE', 'TENTATIVE', 'BUSY', 'AWAY'],
    COLUMNS: {
        'Subject': 'SUMMARY',
        'Start Date': 'DTSTART',
        'Start Time': '',
        'Date Stamp': 'DTSTAMP',
        'End Date': 'DTEND',
        'End Time': '',
        'All Day': '',
        'Description': 'DESCRIPTION',
        'Location': 'LOCATION',
        'UID': 'UID',
        'Busy Status': 'X-MICROSOFT-CDO-BUSYSTATUS'
    },
};

function createICS(name, contents, outfname) {
    var filetext = KEY_WORDS.HEADER;

    filetext += KEY_WORDS.TITLE + name.split('.')[0] + '\n';
    console.log(filetext);
    // Parse csv contents into json
    var json = csvtojson(contents);
    console.log(json);
    // Each element in the json array is an event
    json.forEach(processEvent);

    function processEvent(event) {
        // Don't process is there is no subject
        if (event['Subject'].trim() == "")
            return;

        // Start Event
        filetext += KEY_WORDS.EVENT[0] + '\n';

        const keys = Object.keys(event)
        for (const key of keys) {

            // Ensure that the event exists
            if (KEY_WORDS.COLUMNS[key] && event[key] != '') {

                switch (key) {
                    case 'Start Date':
                        var time = '0:00';

                        // If there is a time, we will add it to the date
                        if (event['Start Time'] && event['Start Time'] != '')
                            time = convert12to24(event['Start Time']);

                        // Check if Event is allday or not
                        if (event['All Day'] && event['All Day'].toLowerCase() == 'true') {
                            var date = new Date(event[key] + ' ' + time);
                            var dtstr = convertDate(date, true);

                            filetext += KEY_WORDS.COLUMNS['Date Stamp'] + ':' + dtstr + 'T000000' + '\n';
                            filetext += KEY_WORDS.COLUMNS[key] + ';VALUE=DATE:' + dtstr + '\n';
                        } else {
                            var date = new Date(event[key] + ' ' + time);
                            var dtstr = convertDate(date, false);

                            filetext += KEY_WORDS.COLUMNS['Date Stamp'] + ':' + dtstr + '\n';
                            filetext += KEY_WORDS.COLUMNS[key] + ':' + dtstr + '\n';
                        }

                        break;

                    case 'End Date':
                        var time = '0:00';

                        // If there is a time, we will add it to the date
                        if (event['End Time'] && event['End Time'] != '')
                            time = convert12to24(event['End Time']);

                        // Check if Event is allday or not
                        if (event['All Day'] && event['All Day'].toLowerCase() == 'true') {
                            var date = new Date(event[key] + ' ' + time);
                            var dtstr = convertDate(date, true);

                            filetext += KEY_WORDS.COLUMNS[key] + ';VALUE=DATE:' + dtstr + '\n';
                        } else {
                            var date = new Date(event[key] + ' ' + time);
                            var dtstr = convertDate(date, false);

                            filetext += KEY_WORDS.COLUMNS[key] + ':' + dtstr + '\n';
                        }

                        break;

                    default:
                        if (event[key] != '')
                            filetext += KEY_WORDS.COLUMNS[key] + ':' + event[key] + '\n';
                }
            }
        }

        // End Event
        filetext += KEY_WORDS.EVENT[1] + '\n';
    }

    function convert12to24(time12h) {
        const time = time12h.slice(0, -2);
        const modifier = time12h.slice(-2).toUpperCase();

        let [hours, minutes] = time.split(':');

        if (hours === '12') {
            hours = '00';
        }

        if (modifier === 'PM') {
            hours = parseInt(hours, 10) + 12;
        }

        return hours + ':' + minutes;
    }

    // Take any ol' date and convert it to the format yyyymmddThhmm00
    function convertDate(date, allday) {
        var pre =
            date.getFullYear().toString() +
            ((date.getMonth() + 1) < 10 ? "0" + (date.getMonth() + 1).toString() : (date.getMonth() + 1).toString()) +
            ((date.getDate() + 1) < 10 ? "0" + date.getDate().toString() : date.getDate().toString());

        var post = (date.getHours() < 10 ? '0' : '') + date.getHours().toString() + (date.getMinutes() < 10 ? '0' : '') + date.getMinutes().toString() + "00";
        if (!allday)
            return pre + "T" + post;
        return pre;
    }

    filetext += KEY_WORDS.FOOTER + '\n';

    gfiles.write(outfname, filetext);

}

// Convert CSV to JSON
function csvtojson(csv) {
    var lines = csv.split("\n");
    var result = [];

    var headers = lines[0].split("|");

    for (var i = 1; i < lines.length; i++) {
        var obj = {};
        var currentline = lines[i].split("|");

        for (var j = 0; j < headers.length; j++) {
            obj[headers[j]] = currentline[j];
        }

        result.push(obj);
    }

    return result; //JSON
}


function main() {
    let opts = gprocs.args("", "infile=calendar.csv,outfile=calendar.ics");
    console.log(opts);
    let csv = gfiles.read(opts.infile);
    console.log(csv);
    createICS("ford calendar", csv, opts.outfile);
}

main();
