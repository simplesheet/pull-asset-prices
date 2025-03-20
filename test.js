function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Asset Prices")
        .addItem("Update Now", "updatePrices")
        .addToUi();
}

function updatePrices() {
    Logger.log("hello from github")
}