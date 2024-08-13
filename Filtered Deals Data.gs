function fetchAndPopulateDeals() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear(); // Clear existing content
  sheet.getRange(1, 1).setValue('Deal Amount');
  sheet.getRange(1, 2).setValue('Owner Name');
  sheet.getRange(1, 3).setValue('Has Object Owner');
  sheet.getRange('1:1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  Logger.log('Headers of the sheet has been set and added!');
  var apiKey = 'Your_Authentication_Token';
  var url = 'https://api.hubapi.com/crm/v3/objects/deals?limit=100&properties=amount,hubspot_owner_id,closedate,pipeline,hs_timestamp,hs_is_closed_won';
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    }
  };
  var moreResults = true;
  var after = null;
  var row = 2;
  Logger.log('Started interacting with HubSpot.');
  Logger.log('Processing your request, please do not close the sheet...');
  while (moreResults) {
    var paginatedUrl = url;
    if (after) {
      paginatedUrl += '&after=' + after;
    }
    var response = UrlFetchApp.fetch(paginatedUrl, options);
    var data = JSON.parse(response.getContentText());
    var deals = data.results;
    for (var i = 0; i < deals.length; i++) {
      var deal = deals[i].properties;
      var pipeline = deal.pipeline;
      var closedwon = deal.hs_is_closed_won
      var timestamp = new Date(deal.closedate);
      var filterDate = new Date('2024-01-01T00:00:00Z');
      if (pipeline === 'default' && closedwon === 'true' && timestamp > filterDate) {
        sheet.getRange(row, 1).setValue(deal.amount || 'N/A');
        var url1 = 'https://api.hubapi.com/owners/v2/owners/' + deal.hubspot_owner_id;
        var options = {
          'method': 'get',
          'headers': {
            'Authorization': 'Bearer Your_Authentication_Token'
          }
        };
        var response1 = UrlFetchApp.fetch(url1, options);
        var data1 = JSON.parse(response1.getContentText());
        var owner = data1.firstName + " " + data1.lastName
        sheet.getRange(row, 2).setValue(owner);
        sheet.getRange(row, 3).setValue(deal.hubspot_owner_id ? 'Yes' : 'No');
        row++;
      }
    }
    if (data.paging && data.paging.next && data.paging.next.after) {
      after = data.paging.next.after;
    } else {
      moreResults = false;
    }
  }
  Logger.log("Deals fetched and populated into Google Sheet.");
  sheet.getDataRange().setBorder(true, true, true, true, true, true);
  Logger.log('Borders have been added.');
  Logger.log('Execution finished. Your data is ready in the sheet. Thank you!');
}
