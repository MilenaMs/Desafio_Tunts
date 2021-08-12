const { google } = require("googleapis");
const keys = require('./keys.json');

//Create client instance for auth
const client = new google.auth.JWT(
    keys.client_email,
    null,
    keys.private_key, ['https://www.googleapis.com/auth/spreadsheets']
);

client.authorize(function(err, tokens) {
    if (err) {
        console.log(err);
        return;
    } else {
        console.log('Connected!')
        gsrun(client);
    }
});

async function gsrun(cl) {

    //Instance of Google Sheets API
    const gsapi = google.sheets({ version: 'v4', auth: cl });
    const opt = {
        spreadsheetId: '1aDehilQcdxmuqGMQLWOXP7zRohbK4TeoBOVDIYPxukw',
        range: 'engenharia_de_software!A4:H27',
    };

    //Read rows from spreadsheet
    let data = await gsapi.spreadsheets.values.get(opt);
    let dataArray = data.data.values;
    let newData = dataArray.map(function(row) {
        (row[7] = 0);
        return row;
    });
    //console.log(newData);


    //Calculate the average
    function Calc(a, b, c) {
        const total = (a + b + c) / 3;
        return total.toFixed(2);
    };

    //Situation for sheet
    dataArray.map(function(row) {

        var lmt = 60 * 0.25;
        var m = Calc(parseFloat(row[3]), parseFloat(row[4]), parseFloat(row[5]));

        var rmd = ((parseFloat(row[3]) + parseFloat(row[4]) + parseFloat(row[5])) % 3);
        if (rmd == 2) {
            m = Math.ceil(m, 2);
        } else if (rmd == 1) {
            m = Math.ceil(m, 1);
        }
        //console.log(m);

        if (parseInt(row[2]) > lmt) {
            row[6] = 'Reprovado por Falta';

        } else if (((m >= 50) && (m < 70)) && (parseInt(row[2]) <= lmt)) {
            var naf = (100 - m).toFixed(2);
            //console.log(naf)
            row[6] = 'Exame Final';
            row[7] = naf;
        } else if ((m < 50) && parseInt(row[2]) <= lmt) {
            row[6] = 'Reprovado por Nota';
        } else {
            row[6] = 'Aprovado';
        }
    });

    //Updating the sheet
    const updateOptions = {
        spreadsheetId: '1aDehilQcdxmuqGMQLWOXP7zRohbK4TeoBOVDIYPxukw',
        range: 'engenharia_de_software!A4:H27',
        valueInputOption: "USER_ENTERED",
        resource: {
            values: newData
        }
    };
    let res = await gsapi.spreadsheets.values.update(updateOptions);

}