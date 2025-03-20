function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Asset Prices")
        .addItem("Update Now", "updatePrices")
        .addToUi();
}

function updatePrices() {
    logger.log("hello")
}