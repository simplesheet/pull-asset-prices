/**
 * 📄 Disclaimer & Usage Policy
 *
 * This script is intended for personal use only and is provided as-is,
 * without warranty of any kind. By using this script, you acknowledge and
 * agree to the following:
 *
 * 🔍 Data Providers
 *
 * Yahoo Finance:
 * - Endpoint: https://query1.finance.yahoo.com
 * - This uses internal Yahoo endpoints not officially documented or supported.
 * - Use is subject to Yahoo's Terms of Service:
 *   https://legal.yahoo.com/us/en/yahoo/terms/otos/index.html
 * - This script is not affiliated with or endorsed by Yahoo.
 *
 * KuCoin:
 * - Endpoint: https://api.kucoin.com/api/v1/market/allTickers
 * - Publicly accessible and does not require authentication.
 * - Governed by KuCoin's Terms of Use:
 *   https://www.kucoin.com/legal/terms-of-use
 * - This script is not affiliated with or endorsed by KuCoin.
 *
 * Gold-API:
 * - Website: https://gold-api.com
 * - Subject to their Terms of Use: https://gold-api.com/terms
 * - This script is not affiliated with or endorsed by Gold-API.
 *
 * 📌 Usage Restrictions
 * - Do not use this script for commercial purposes unless licensed.
 * - Do not redistribute or expose raw data in public apps or products.
 * - Respect all API rate limits and fair use guidelines.
 *
 * 💬 Transparency
 * This script was created for educational and personal portfolio tracking use.
 * If you are the owner of one of the services mentioned above and have concerns,
 * please contact the maintainer to discuss proper usage or removal.
 *
 * ☕ Support
 * This script is free to use. If you'd like to support development,
 * you can donate here: https://your-ko-fi-or-donation-link
 */

const VERSION = "1.0.0";
const ASSETS_SHEET_NAME = "Asset Prices";
const UPDATE_RESULTS_RANGE = "update_results";
const LAST_UPDATED_RANGE = "prices_last_updated";

function init() {
  Logger.log("Initializing project...")
  // Delete any existing 'updatePrices' triggers
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "updatePrices") {
      Logger.log("Deleting existing 'updatePrices' trigger");
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Add a new trigger that updates prices every hour
  Logger.log("Adding new 'updatePrices' trigger (update prices every hour)");
  ScriptApp.newTrigger("updatePrices")
    .timeBased()
    .everyHours(1)
    .create();

  // Build the 'Asset Prices' sheet
  rebuildAssetPriceSheet();
  Logger.log("Project initialized!")
}


function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Pull Asset Prices")
    .addItem("Pull Prices Now", "updatePrices")
    .addItem("Price Variable - Usage Example", "displayPriceVariableExampleModal")
    .addItem("Check For Updates", "displayVersionCheckModal")
    .addToUi();
}


function displayVersionCheckModal() {
  const localVersion = VERSION;
  const remoteVersion = getVersionFromGithubScript();
  
  if (remoteVersion) {
    var updateAvailable = false;

    const localVersionParts = localVersion.split(/\./);
    const remoteVersionParts = remoteVersion.split(/\./);
    for (var i = 0; i < localVersionParts.length; i++) {
      if (remoteVersionParts[i] > localVersionParts[i]) {
        updateAvailable = true;
      }
    }

    if (updateAvailable) {
      var html = `<h4>An update is available! 🤖</h4>
      <p>Local Version: ` + localVersion + `<br>
      Remote Version: ` + remoteVersion + `</p>
      <p>Updates might include bug fixes or new features. If the tool is working well for you, there's no reason you need to update.</p>
      View the <a href="https://simplesheet.github.io/docs/tools/pull-asset-prices/update.html" target="_blank" rel="noopener">Update Instructions</a> if you would like to update.`;

    } else {
      var html = `<h4>You're all up to date! 🎉</h4>
      <p>Your Version: ` + localVersion + `</p>`;
    }

  } else {
    var html = `<h4>There was a issue checking for updates 🤔</h4>
    <p>Your Version: ` + localVersion + `</p>
    Please try again later...`;
  }

  var htmlObject = HtmlService.createHtmlOutput(html)
                              .setWidth(500)
                              .setHeight(210);
  SpreadsheetApp.getUi().showModalDialog(htmlObject, "Check For Updates");
}


function getVersionFromGithubScript() {
  const url = 'https://raw.githubusercontent.com/simplesheet/pull-asset-prices/refs/heads/main/pull_asset_prices.gs';
  const response = UrlFetchApp.fetch(url);
  const scriptContent = response.getContentText();

  const versionMatch = scriptContent.match(/VERSION\s*=\s*["']([^"']+)["']/);
  if (versionMatch && versionMatch[1]) {
    Logger.log('Remote VERSION: ' + versionMatch[1]);
    return versionMatch[1];
  } else {
    Logger.log('VERSION not found.');
    return null;
  }
}


function displayPriceVariableExampleModal() {
  var spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var html = `<p>Using <strong>BTC</strong> as an example...</p>
<h4 style="margin-bottom: 2px;">To use the price of BTC on another tab in THIS spreadsheet</h4>
<table border="1" style="border-collapse: collapse; width: 100%;">
    <tbody>
        <tr>
            <td style="border: 1px solid #aaa; padding: 8px; text-align: left; background-color: #AFDDC0">Enter the
                following formula into a cell:<br> <strong>=BTC_Price</strong></td>
        </tr>
    </tbody>
</table>

<h4 style="margin-bottom: 2px;">To use the price of BTC in a DIFFERENT spreadsheet</h4>
<table border=" 1" style="border-collapse: collapse; width: 100%;">
    <tbody>
        <tr>
            <td style="border: 1px solid #aaa; padding: 8px; text-align: left; background-color: #AFDDC0">Enter the
                following formula into a cell: <strong>=IMPORTRANGE("` + spreadsheetUrl + `", "BTC_Price")</strong></td>
        </tr>
    </tbody>
</table>`;
  var htmlObject = HtmlService.createHtmlOutput(html)
                              .setWidth(600)
                              .setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(htmlObject, "Price Variable Example");
}


function updatePrices() {
  checkBannerImage();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var assetPricesSheet = spreadsheet.getSheetByName(ASSETS_SHEET_NAME);

  if (!assetPricesSheet) {
    init();
    var assetPricesSheet = spreadsheet.getSheetByName(ASSETS_SHEET_NAME);
  }

  // Reset Update Results (logging) output
  spreadsheet.getRange(UPDATE_RESULTS_RANGE).setValue("");
  

  // Get Asset Tickers
  const stocks = spreadsheet.getRange('stock_tickers');
  const crypto = spreadsheet.getRange('crypto_tickers');
  const metals = spreadsheet.getRange('metal_names');
  const assets = [stocks, crypto, metals]

  try {
    // Get all the crypto prices
    var cryptoApiData = getCryptoData()

  } catch (error) {
    addLogEntry(spreadsheet, error)
    var cryptoApiData = []
  }
  
  var processedTickers = [];
  // Get tickers for each asset class 
  // pull prices for each and set a Named Range 
  assets.forEach((asset, index) => {
    var numRows = asset.getNumRows();
    var tickerCol = asset.getColumn(); 
    var priceCol = tickerCol + 1;
    var variableCol = priceCol + 1; 
    var startRow = asset.getRow(); 
    
    for (var i = 0; i < numRows; i++) {
      var price = 0;
      var currentTickerCell = assetPricesSheet.getRange(startRow + i, tickerCol);
      var currentTickerValue = currentTickerCell.getValue();
      var currentPriceCell = assetPricesSheet.getRange(startRow + i, priceCol);
      var currentPriceVariableCell = assetPricesSheet.getRange(startRow + i, variableCol);

      if (currentTickerValue) {

        // Set Ticker value to UPPERCASE if needed
        if (currentTickerValue != currentTickerValue.toUpperCase()) {
          const currentTickerValueUpper = currentTickerValue.toUpperCase();
          addLogEntry(spreadsheet, "Changed Ticker '" + currentTickerValue + "' to Uppercase: " + currentTickerValueUpper);
          currentTickerValue = currentTickerValueUpper;
          currentTickerCell.setValue(currentTickerValueUpper);  
        };

        // Make sure the Ticker isn't a duplicate entry
        if (!processedTickers.includes(currentTickerValue)) {
          processedTickers.push(currentTickerValue);

          try {
              // Stocks
              if (index == 0) {
                price = getStockPrice(currentTickerValue)
                if (price === null) {
                  price = 0;
                  addLogEntry(spreadsheet, "[ERROR] No price data found for Stock/ETF asset: " + currentTickerValue);
                }
              }

              // Crypto
              else if (index == 1) {
                // Adjust ticker if needed
                if (currentTickerValue === 'GALA') {
                  var kucoinTicker = 'GALAX'
                } else {
                  var kucoinTicker = currentTickerValue
                }
                price = locateCryptoPrice(kucoinTicker, cryptoApiData)

                if (price === 0) {
                  addLogEntry(spreadsheet, "[ERROR] No price data found for Crypto asset: " + currentTickerValue);
                }
                
              }

              // Metals
              else if (index == 2) {
                price = getMetalPrice(currentTickerValue);

                if (price === 0) {
                  addLogEntry(spreadsheet, "[ERROR] No price data found for Metal asset: " + currentTickerValue);
                }
              }

          } catch (error) {
            addLogEntry(spreadsheet, error);
          }

          // Convert the price to a number format if needed
          // This is important for Localization formatting
          if (typeof price === 'string') {
            price = Number(price);
          }

          // Make sure the number conversion worked before setting the value
          if (typeof price === 'number' && !isNaN(price)) {
            // Set price Value
            currentPriceCell.setValue(price);
            addLogEntry(spreadsheet, currentTickerValue + " price updated to $" + price);
          } else {
            addLogEntry(spreadsheet, currentTickerValue + " Invalid price format: '" + typeof price + "'. Failed to convert to number-type.");
            price = 0;
          }

          if (price != 0) {
            // Build Named Range 'Name'
            var rangeName = currentTickerValue + "_Price";
            
            try {
              var existingNamedRange = assetPricesSheet.getRange(rangeName);
            } catch (e) {
              var existingNamedRange = null;
            }
            if (!existingNamedRange) {
              // This is a new named range
              addLogEntry(spreadsheet, "New named range '" + rangeName + "' set for cell " + currentPriceCell.getA1Notation());
            }

            // Set named range
            spreadsheet.setNamedRange(rangeName, currentPriceCell);
            currentPriceVariableCell.setValue(rangeName);
          
          } else {
            currentPriceVariableCell.setValue("ERROR");
          }

        } else {
          // Ticker is a duplicate 
          addLogEntry(spreadsheet, "[WARN] The ticker '" + currentTickerValue + "' at " + currentTickerCell.getA1Notation() + " is a duplicate!");
          currentPriceCell.setValue(price);
          currentPriceVariableCell.setValue('DUPLICATE')
        }

      } else {
        // No Ticker value in cell
        // Check if the price cell has a value
        // If it does then that means the corresponding ticker was removed
        if (currentPriceCell.getValue() || currentPriceCell.getValue() == '0') {
          addLogEntry(spreadsheet, "A ticker value was removed from " + currentTickerCell.getA1Notation() + " since last run");
          currentPriceCell.setValue('');
          currentPriceVariableCell.setValue('');
          addLogEntry(spreadsheet, "Orphaned price value removed from " + currentPriceCell.getA1Notation()); 
          addLogEntry(spreadsheet, "Orphaned price variable name removed from " + currentPriceVariableCell.getA1Notation());

          const namedRanges = spreadsheet.getNamedRanges();

          for (var rangeIndex = 0; rangeIndex < namedRanges.length; rangeIndex++) {
              if (namedRanges[rangeIndex].getRange().getA1Notation() === currentPriceCell.getA1Notation()) {
                  addLogEntry(spreadsheet, "Removed Named Range " + namedRanges[rangeIndex].getName() + " at " + currentPriceCell.getA1Notation());
                  namedRanges[rangeIndex].remove();
              }
          }
        }
      }
    }
  })

  // Set the 'Prices Last Updated' date and time
  const unixdate = new Date()
  const timeZone = Session.getScriptTimeZone()
  const formattedDate = Utilities.formatDate(unixdate, timeZone, 'E MMM dd yyyy')
  const time = unixdate.toLocaleTimeString("en-US")
  spreadsheet.getRange(LAST_UPDATED_RANGE).setValue(formattedDate + " " + time)

}


function addLogEntry(spreadsheet, newLogEntry) {
  var logRange = spreadsheet.getRange(UPDATE_RESULTS_RANGE);
  var currentLogValue = logRange.getValue();

  // Use regex to find all occurences of \n to count the lines
  var logLines = (currentLogValue.match(/\n/g) || []).length;
  var currentLogLine = logLines + 1;

  var newLogValue = currentLogValue + currentLogLine + ". " + newLogEntry + "\n";
  
  logRange.setValue(newLogValue);
}



///////////////////////////////
//       API Functions       //
///////////////////////////////  

// https://finance.yahoo.com/
// This is an 'unofficial' API so it isn't well documented and could change
// Limit: 2,000 per hour. 5-10 per second.
// Each updatePrices() execution uses 1 request for each stock in the list
function getStockPrice(ticker) {
  try {
    var url = "https://query1.finance.yahoo.com/v8/finance/chart/" + ticker;
    var response = UrlFetchApp.fetch(url);
    
    if (response.getResponseCode() == 200) {
      var json = JSON.parse(response.getContentText());

      var price = json.chart.result[0].meta.regularMarketPrice;
      if (price) {
        return price;
      } else {
        return null;
      }

    } else {
      console.error("Error fetching stock price for " + ticker + ": " + response.getResponseCode());
      return null;
    }

  } catch (error) {
    console.error("Error fetching stock price for " + ticker + ": " + error);
    return null;
  }
  
}


// Crypto
// https://www.kucoin.com/trade/BTC-USDT
// This tool is not affiliated with or endorsed by KuCoin. All data belongs to KuCoin. Use at your own discretion.
// Limit: 20 requests per second, 1,200 requests per minute
// Each updatePrices() execution uses 1 request
function getCryptoData() {
  var apiURL = 'https://api.kucoin.com/api/v1/market/allTickers';
  var response = UrlFetchApp.fetch(apiURL);
  if (response.getResponseCode() == 200) {
    var json = JSON.parse(response.getContentText())
    // Return array of all crypto ticker objects
    return json['data']['ticker'];

  } else {
    console.error("Error fetching crypto data: " + response.getResponseCode());
    return null;
  }
}

function locateCryptoPrice(ticker, cryptoData) {
  var price = 0;
  // Format the ticker into the KuCoin pair
  kucoinPair = ticker + '-USDT'

  for (var index in cryptoData) {
    currentCryptoData = cryptoData[index];
    if (currentCryptoData['symbol'] === kucoinPair) {
      var cryptoPriceRange = ticker.toLowerCase() + 'Price';
      price = currentCryptoData['last'];
      break;
    }
  }
  return price;
}

// Metals
// https://gold-api.com/docs
// Limit: No limits
function getMetalPrice(metal) {
    metalSymbols = {
      'GOLD': 'XAU',
      'SILVER': 'XAG',
      'PLATINUM': 'XPT'
    }
    var symbol = metalSymbols[metal.toUpperCase()];
    if (symbol) {
      var url = "https://api.gold-api.com/price/" + symbol;
      var response = UrlFetchApp.fetch(url);
    
      if (response.getResponseCode() == 200) {
          var json = JSON.parse(response.getContentText());
          var price = json.price;
          return price;

      } else {
          console.error("Error fetching metal data for " + metal + ": " + response.getResponseCode());
          return null;
      }

    } else {
      console.error("The passed in metal (" + metal + ") has no matching symbol");
      return null;
    }
}




///////////////////////////////
//  Build Asset Price Sheet  //
/////////////////////////////// 

function rebuildAssetPriceSheet() {
  Logger.log("Rebuilding Asset Price Sheet...")
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  assetsSheet = spreadsheet.getSheetByName(ASSETS_SHEET_NAME);
  if (!assetsSheet) {
    Logger.log("Building price list")
    buildAssetPriceSheet();
    Logger.log("'Asset Prices' sheet built")
  } else {
    Logger.log("Sheet already exists")
    Logger.log("Deleting and rebuilding...")
    spreadsheet.deleteSheet(assetsSheet);
    buildAssetPriceSheet();
    Logger.log("'Asset Prices' sheet rebuilt")
  }
}


function buildAssetPriceSheet() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.insertSheet(ASSETS_SHEET_NAME);

    //
    // Banner and Info area
    //
    var bannerStartRow = 1;
    var bannerEndRow = bannerStartRow + 4;
    var usefulLinksStartColumn = "D";
    var usefulLinksEndColumn = "F";
    var lastUpdatedStartColumn = "H";
    var lastUpdatedEndColumn = "J";
    var usefulLinksTitleStartRow = bannerEndRow + 2;
    var usefulLinksTitleEndRow = usefulLinksTitleStartRow + 1;
    var instructionsLinkRow = usefulLinksTitleEndRow + 1;
    var featureRequestLinkRow = instructionsLinkRow + 1;
    var donateLinkRow = featureRequestLinkRow + 1;

    // Banner photo and background color
    var bannerRange = sheet.getRange("B" + bannerStartRow + ":L" + bannerEndRow);
    bannerRange.merge();
    bannerRange.setBackground("#8FC2A1");
    setBannerPhoto();

    // Useful Links
    var usefulLinksTitleRange = sheet.getRange(usefulLinksStartColumn + usefulLinksTitleStartRow + ":" + usefulLinksEndColumn + usefulLinksTitleEndRow);
    usefulLinksTitleRange.merge();
    usefulLinksTitleRange.setValue("Useful Links");
    usefulLinksTitleRange.setBackground("#8FC2A1")
                         .setFontColor("#434343")
                         .setHorizontalAlignment("center")
                         .setVerticalAlignment("middle")
                         .setFontFamily("Comfortaa")
                         .setFontSize(16)
                         .setFontWeight("bold");

    // Instructions Link
    var instructionsLinkRange = sheet.getRange(usefulLinksStartColumn + instructionsLinkRow + ":" + usefulLinksEndColumn + instructionsLinkRow);
    instructionsLinkRange.merge()
    const instructionsRichText = SpreadsheetApp.newRichTextValue()
      .setText("Full Instructions")
      .setLinkUrl("https://simplesheet.github.io/docs/tools/pull-asset-prices/instructions.html")
      .build();

    instructionsLinkRange.setRichTextValue(instructionsRichText)
                         .setHorizontalAlignment("center")
                         .setVerticalAlignment("middle")
                         .setFontColor("black")
                         .setFontWeight("bold"); 
  
    // Feature Request Link
    var featureRequestLinkRange = sheet.getRange(usefulLinksStartColumn + featureRequestLinkRow + ":" + usefulLinksEndColumn + featureRequestLinkRow);
    featureRequestLinkRange.merge()
    const featureRequestRichText = SpreadsheetApp.newRichTextValue()
      .setText("Request a feature or report an issue")
      .setLinkUrl("https://docs.google.com/forms/d/e/1FAIpQLSce9-dAMIRSN--Opz6fI4-sTJrvzK_IRTJAGiiL6SsmF4pSpQ/viewform?usp=header")
      .build();

    featureRequestLinkRange.setRichTextValue(featureRequestRichText)
                          .setHorizontalAlignment("center")
                          .setVerticalAlignment("middle")
                          .setFontColor("black")
                          .setFontWeight("bold"); 

    // Feature Request Link
    var donateLinkRange = sheet.getRange(usefulLinksStartColumn + donateLinkRow + ":" + usefulLinksEndColumn + donateLinkRow);
    donateLinkRange.merge()
    const donateRichText = SpreadsheetApp.newRichTextValue()
      .setText("Want to support development? Donate here.")
      .setLinkUrl("https://ko-fi.com/simplesheet")
      .build();

    donateLinkRange.setRichTextValue(donateRichText)
                          .setHorizontalAlignment("center")
                          .setVerticalAlignment("middle")
                          .setFontColor("black")
                          .setFontWeight("bold"); 

    // Set useful links border
    var usefulLinksRange = sheet.getRange(usefulLinksStartColumn + usefulLinksTitleStartRow + ":" + usefulLinksEndColumn + donateLinkRow);
    usefulLinksRange.setBorder(
      true, true, true, true, false, false, 
      "#8FC2A1",
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );

    // Prices Last Updated
    var lastUpdateTitleRange = sheet.getRange(lastUpdatedStartColumn + usefulLinksTitleStartRow + ":" + lastUpdatedEndColumn + usefulLinksTitleEndRow);
    lastUpdateTitleRange.merge();
    lastUpdateTitleRange.setValue("Prices Last Updated");
    lastUpdateTitleRange.setBackground("#8FC2A1")
                         .setFontColor("#434343")
                         .setHorizontalAlignment("center")
                         .setVerticalAlignment("middle")
                         .setFontFamily("Comfortaa")
                         .setFontSize(16)
                         .setFontWeight("bold");

    var lastUpdatedTitleRange = sheet.getRange(lastUpdatedStartColumn + instructionsLinkRow + ":" + lastUpdatedEndColumn + donateLinkRow);
    lastUpdatedTitleRange.merge()
    lastUpdatedTitleRange.setValue("Never!")
                         .setHorizontalAlignment("center")
                         .setVerticalAlignment("middle")
                         .setFontSize(14);

    spreadsheet.setNamedRange(LAST_UPDATED_RANGE, lastUpdatedTitleRange);

    // Set useful links border
    var lastUpdatedRange = sheet.getRange(lastUpdatedStartColumn + usefulLinksTitleStartRow + ":" + lastUpdatedEndColumn + donateLinkRow);
    lastUpdatedRange.setBorder(
      true, true, true, true, false, false, 
      "#8FC2A1",
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );

    //
    // ASSETS AREA
    //
    var assetsTitleStartRow = 13;
    var assetsTitleEndRow = assetsTitleStartRow + 3;
    var assetSectionsTitleStartRow = assetsTitleEndRow + 1;
    var assetSectionsTitleEndRow = assetSectionsTitleStartRow + 1;
    var assetsHeaderRow = assetSectionsTitleEndRow + 1;
    var assetsDataStartRow = assetsHeaderRow + 1;
    var stockEndRow = assetsDataStartRow + 49;
    var cryptoEndRow = assetsDataStartRow + 49;
    var metalEndRow = assetsDataStartRow + 1;

    var assetsTitleRange = sheet.getRange("B" + assetsTitleStartRow + ":L" + assetsTitleEndRow);
    assetsTitleRange.merge();
    assetsTitleRange.setBackground("#434343")
                    .setFontColor("white")
                    .setHorizontalAlignment("center")
                    .setVerticalAlignment("middle");

    
    var text = "ASSETS\nAdd/remove your desired Stock & Crypto tickers in the green columns - then from the menu bar select 'Pull Asset Prices > Pull Prices Now'.\nPrices automatically update hourly but they can be updated at any time from the menu.";

    // Create a RichTextValueBuilder
    var richTextBuilder = SpreadsheetApp.newRichTextValue().setText(text);

    // Apply font size styles using a TextStyleBuilder
    var largeFontStyle = SpreadsheetApp.newTextStyle().setFontSize(18).build();
    var mediumFontStyle = SpreadsheetApp.newTextStyle().setFontSize(12).build();

    // Apply styles to different portions of the text
    richTextBuilder.setTextStyle(0, 6, largeFontStyle);  // "ASSETS" in 18px 
    richTextBuilder.setTextStyle(7, text.length, mediumFontStyle);  // Rest of text in 12px

    // Build the formatted rich text
    var assetsTitleText = richTextBuilder.build();

    // Set the formatted text in the merged cell range
    assetsTitleRange.setRichTextValue(assetsTitleText);

    assetsTitleRange.setBorder(
      true, true, true, true,    // Top, left, bottom, right borders
      false,                     // Inner horizontal borders
      false                      // Inner vertical borders
    );


  //
  // ASSET AREAS
  //
  var stocksTitle = "STOCK/ETF";
  var cryptoTitle = "CRYPTO";
  var metalTitle = "METAL";
  var assetSections = [
    {
      "titleText": stocksTitle,
      "tickerNamedRange": "stock_tickers",
      "titleRange": "B" + assetSectionsTitleStartRow + ":D" + assetSectionsTitleEndRow,
      "tickerHeaderRange": "B" + assetsHeaderRow,
      "priceHeaderRange": "C" + assetsHeaderRow,
      "priceVariableHeaderRange": "D" + assetsHeaderRow,
      "tickerRange": "B" + assetsDataStartRow + ":B" + stockEndRow,
      "priceRange": "C" + assetsDataStartRow + ":C" + stockEndRow,
      "priceVariableRange": "D" + assetsDataStartRow + ":D" + stockEndRow
    },
    {
      "titleText": cryptoTitle,
      "tickerNamedRange": "crypto_tickers",
      "titleRange": "F" + assetSectionsTitleStartRow + ":H" + assetSectionsTitleEndRow,
      "tickerHeaderRange": "F" + assetsHeaderRow,
      "priceHeaderRange": "G" + assetsHeaderRow,
      "priceVariableHeaderRange": "H" + assetsHeaderRow,
      "tickerRange": "F" + assetsDataStartRow + ":F" + cryptoEndRow,
      "priceRange": "G" + assetsDataStartRow + ":G" + cryptoEndRow,
      "priceVariableRange": "H" + assetsDataStartRow + ":H" + cryptoEndRow
    },
    {
      "titleText": metalTitle,
      "tickerNamedRange": "metal_names",
      "titleRange": "J" + assetSectionsTitleStartRow + ":L" + assetSectionsTitleEndRow,
      "tickerHeaderRange": "J" + assetsHeaderRow,
      "priceHeaderRange": "K" + assetsHeaderRow,
      "priceVariableHeaderRange": "L" + assetsHeaderRow,
      "tickerRange": "J" + assetsDataStartRow + ":J" + metalEndRow,
      "priceRange": "K" + assetsDataStartRow + ":K" + metalEndRow,
      "priceVariableRange": "L" + assetsDataStartRow + ":L" + metalEndRow
    }
  ];

  var editableRanges = [];
  var formatRules = [];
  assetSections.forEach(section => {
    // Logger.log("titleRange: " + section.titleRange)
    var titleRange = sheet.getRange(section.titleRange);
    titleRange.merge();
    titleRange.setBackground("#666666")
              .setFontColor("white")
              .setFontSize(16)
              .setHorizontalAlignment("center")
              .setVerticalAlignment("middle");

    titleRange.setValue(section.titleText);

    titleRange.setBorder(true, true, true, true, false, false);

    //
    // Set the Asset Section headers and columns
    //
    var tickerHeader = "Ticker";
    var priceHeader = "Price";
    var priceVariableHeader = "Price Variable";
    var columnns = [
      {
        "titleText": tickerHeader,
        "headerRange": section.tickerHeaderRange,
        "columnRange": section.tickerRange,
        "namedRange":  section.tickerNamedRange
      },
      {
        "titleText": priceHeader,
        "headerRange": section.priceHeaderRange,
        "columnRange": section.priceRange
      },
      {
        "titleText": priceVariableHeader,
        "headerRange": section.priceVariableHeaderRange,
        "columnRange": section.priceVariableRange
      }
    ]

    

    // Set each column for the Asset Section
    // Ticker, Price, Price Variable
    columnns.forEach(column => {
      // Logger.log("headerRange: " + column.headerRange)
      var headerRange = sheet.getRange(column.headerRange);
      headerRange.setFontWeight("bold")
        .setFontSize(10)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setValue(column.titleText)
        .setBorder(true, true, true, true, false, false);

      // Set the columns
      var columnRange = sheet.getRange(column.columnRange);
      columnRange.setBorder(true, true, true, true, false, false);

      var titleText = column.titleText;
      // 'Ticker' header
      if (titleText === tickerHeader) {
        if (section.titleText != metalTitle) {
          // METAL tickers are hard coded, so only allow Stock and Crypto to be edited
          editableRanges.push(columnRange);
          columnRange.setBackground("#AFDDC0");
        }

        // Apply the named range used for pulling in the Tickers for price updates
        spreadsheet.setNamedRange(column.namedRange, columnRange);


      // 'Price' header
      } else if (titleText === priceHeader) {
        // Set format to Dollars
        columnRange.setNumberFormat("$#,##0.00");

        // Format text to red when price is $0.00
        var formatRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(0)
          .setFontColor("#FF0000")
          .setRanges([columnRange])
          .build();

        formatRules.push(formatRule)


      // 'Price Variable' header
      } else if (titleText === priceVariableHeader) {
        // Format text to red when value is DUPLICATE or ERROR
        var columnLetter = columnRange.getA1Notation().replace(/\d/g, "");
        var formatRule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=OR($' + columnLetter + assetsDataStartRow + '="DUPLICATE", \
                                     $' + columnLetter + assetsDataStartRow + '="ERROR" )')
          .setFontColor("#FF0000")
          .setBold(true)
          .setRanges([columnRange])
          .build();

        formatRules.push(formatRule)
      }

    }); // End columnns.forEach

    // Add a starter ticker value for Stock and Crypto
    sheet.getRange("B" + assetsDataStartRow).setValue('TSLA');
    sheet.getRange("F" + assetsDataStartRow).setValue('BTC');

    // Set Metal Names/Tickers
    if (section.titleText === metalTitle) {
      sheet.getRange("J" + assetsDataStartRow)
        .setBackground("#D9D9D9")
        .setValue("GOLD");
      sheet.getRange("J" + metalEndRow)
        .setBackground("#D9D9D9")
        .setValue("SILVER");
    }


  }); // End assetSections.forEach

  // Apply conditional format rules and protection
  sheet.setConditionalFormatRules(formatRules);
  var protection = sheet.protect().setDescription("Only Allow Ticker Edits");
  protection.setUnprotectedRanges(editableRanges);
  protection.setWarningOnly(true);

  //
  // Update Results (logging) section
  //
  var updateResultsStartColumn = 'N';
  var updateResultsEndColumn = 'Q';
  var updateResultsStartRow = 2;
  var updateResultsDescStart = updateResultsStartRow + 1;
  var updateResultsDescEnd = updateResultsDescStart + 4;
  var updateResultsLogStart = updateResultsDescEnd + 1;
  var updateResultsLogEnd =  updateResultsLogStart + 124;

  // Update Results Title 
  var updateResultsTitleRange = sheet.getRange(updateResultsStartColumn + updateResultsStartRow + ":" + updateResultsEndColumn + updateResultsStartRow);
  updateResultsTitleRange.merge()
    .setBorder(true, true, true, true, false, false)
    .setValue("Last Update Results")
    .setFontSize(12)
    .setFontColor("white")
    .setBackground("#666666")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Update Results Description
  var zeroPriceFontStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("red").build();
  var updateResultsRichText = SpreadsheetApp.newRichTextValue()
    .setText("The results from the last update are logged here.\n\nIf you get a price of $0.00 for one of your listed assets:\n - It means an issue occured.\n - Review the messages below for an explanation.")
    .setTextStyle(72, 78, zeroPriceFontStyle)
    .build();
  var updateResultsDescRange = sheet.getRange(updateResultsStartColumn + updateResultsDescStart + ":" + updateResultsEndColumn + updateResultsDescEnd);
  updateResultsDescRange.merge()
    .setBorder(true, true, true, true, false, false)
    .setVerticalAlignment("top")
    .setWrap(true)
    .setRichTextValue(updateResultsRichText);

  // Update Results Logs
  var updateResultsLogRange = sheet.getRange(updateResultsStartColumn + updateResultsLogStart + ":" + updateResultsEndColumn + updateResultsLogEnd);
  updateResultsLogRange.merge()
    .setBorder(true, true, true, true, false, false)
    .setBackground("#D9D9D9")
    .setVerticalAlignment("top")
    .setWrap(true)
    .setValue("1. Asset Prices sheet initialized! Price update logs will be written here.");

  // Set the named range for writing logs to the Update Results area
  spreadsheet.setNamedRange(UPDATE_RESULTS_RANGE, updateResultsLogRange)

}


function setBannerPhoto() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ASSETS_SHEET_NAME);

  const imageUrl = "https://raw.githubusercontent.com/simplesheet/pull-asset-prices/refs/heads/main/pull-asset-prices-banner.png";
  
  const startColumn = 2; // Column B
  const endColumn = 12; // Column L
  const targetRow = 1;

  // Fetch the image
  const blob = UrlFetchApp.fetch(imageUrl).getBlob();
  const image = sheet.insertImage(blob, startColumn, targetRow);

  // Calculate total pixel width from B to L
  let totalWidth = 0;
  for (let col = startColumn; col <= endColumn; col++) {
    totalWidth += sheet.getColumnWidth(col);
  }

  // Resize image to fit across columns B to L
  const originalHeight = image.getHeight();
  const originalWidth = image.getWidth();
  const aspectRatio = originalHeight / originalWidth;

  image.setWidth(totalWidth);
  image.setHeight(totalWidth * aspectRatio);
}


function checkBannerImage() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ASSETS_SHEET_NAME);
  var images = sheet.getImages();

  if (images.length === 0) {
    Logger.log("No floating images found.");
    setBannerPhoto();
  }
}

