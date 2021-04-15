var sh2=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coins");

function getCryptoPrice() {

  var finished = false;
  var index = 2;
  while (!finished) {
    if(sh2.getRange(index, 1).getValue().toString().trim().valueOf() == 0){
      finished = true;
      break;
    }
    fetchData(index)
    index++

  }

}

function fetchData(index){

  var sh1=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ApiKey");
  var apiKey=sh1.getRange(1, 2).getValue();

  var url="https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol="+ sh2.getRange(index, 1).getValue().toString().trim();

  var requestOptions = {
    method: 'GET',
    uri: 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest',
    qs: {
      start: 1,
      limit: 5000,
      convert: 'USD'
    },
    headers: {
      'X-CMC_PRO_API_KEY': apiKey
    },
    json: true,
    gzip: true,
    muteHttpExceptions: true
  };

  var httpRequest = UrlFetchApp.fetch(url, requestOptions);
  if(httpRequest.getResponseCode() == 200){
    var getContext= httpRequest.getContentText();
    var coinId = sh2.getRange(index, 1).getValue().toString().trim();
    var parseData=JSON.parse(getContext);
    //console.log(parseData.data[coinId].quote.USD.price.toFixed(4));
    sh2.getRange(index, 4).setValue(parseData.data[coinId].quote.USD.price)
  }
}
