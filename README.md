# Training_Participants_Report

let loadedResolve, reportLoaded = new Promise((res, rej) => { loadedResolve = res; });
let renderedResolve, reportRendered = new Promise((res, rej) => { renderedResolve = res; });

// Get models. models contains enums that can be used.
models = window['powerbi-client'].models;

// Embed a Power BI report in the given HTML element with the given configurations
// Read more about how to embed a Power BI report in your application here: https://go.microsoft.com/fwlink/?linkid=2153590
function embedPowerBIReport() {
    /*-----------------------------------------------------------------------------------+
    |    Don't change these values here: access token, embed URL and report ID.          | 
    |    To make changes to these values:                                                | 
    |    1. Save any other code changes to a text editor, as these will be lost.         |
    |    2. Select 'Start over' from the ribbon.                                         |
    |    3. Select a report or use an embed token.                                       |
    +-----------------------------------------------------------------------------------*/
    // Read embed application token
    let accessToken = EMBED_ACCESS_TOKEN;

    // Read embed URL
    let embedUrl = EMBED_URL;

    // Read report Id
    let embedReportId = REPORT_ID;

    // Read embed type from radio
    let tokenType = TOKEN_TYPE;

    // We give All permissions to demonstrate switching between View and Edit mode and saving report.
    let permissions = models.Permissions.All;

    // Create the embed configuration object for the report
    // For more information see https://go.microsoft.com/fwlink/?linkid=2153590
    let config = {
        type: 'report',
        tokenType: tokenType == '0' ? models.TokenType.Aad : models.TokenType.Embed,
        accessToken: accessToken,
        embedUrl: embedUrl,
        id: embedReportId,
        permissions: permissions,
        settings: {
            panes: {
                filters: {
                    visible: true
                },
                pageNavigation: {
                    visible: true
                }
            },
            bars: {
                statusBar: {
                    visible: true
                }
            }
        }
    };

    // Get a reference to the embedded report HTML element
    let embedContainer = $('#embedContainer')[0];

    // Embed the report and display it within the div container.
    report = powerbi.embed(embedContainer, config);

    // report.off removes all event handlers for a specific event
    report.off("loaded");

    // report.on will add an event handler
    report.on("loaded", function () {
        loadedResolve();
        report.off("loaded");
    });

    // report.off removes all event handlers for a specific event
    report.off("error");

    report.on("error", function (event) {
        console.log(event.detail);
    });

    // report.off removes all event handlers for a specific event
    report.off("rendered");

    // report.on will add an event handler
    report.on("rendered", function () {
        renderedResolve();
        report.off("rendered");
    });
}

embedPowerBIReport();
await reportLoaded;

// Insert here the code you want to run after the report is loaded

await reportRendered;

// Insert here the code you want to run after the report is rendered

