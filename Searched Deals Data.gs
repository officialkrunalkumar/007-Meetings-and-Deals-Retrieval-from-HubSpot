function fetchAndPopulateDeals() {
  var inputs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visual');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Deals');
  sheet.clear(); // Clear existing content
  sheet.getRange(1, 1).setValue('Deal Amount');
  sheet.getRange(1, 2).setValue('Account Executive');
  sheet.getRange(1, 3).setValue('Has Object Owner');
  sheet.getRange(1, 4).setValue('Deal Close Date (YYYY-MM-DD)');
  sheet.getRange(1, 5).setValue('Deal Close Time (HH:MM:SS)');
  sheet.getRange(1, 6).setValue('Deal Type');
  sheet.getRange(1, 7).setValue('Pipeline');
  sheet.getRange(1, 8).setValue('Is Closed Won?')
  sheet.getRange('1:1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  Logger.log('Headers of the sheet has been set and added!');
  var fdate = inputs.getRange('B3').getValue();
  var date = new Date(fdate);
  var apiKey = 'Your_Authentication_Token';
  var url = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  var options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify({
      "filterGroups": [
        {
          "filters": [
            {
              "propertyName": "pipeline",
              "operator": "EQ",
              "value": "default"
            },
            {
              "propertyName": "dealtype",
              "operator": "EQ",
              "value": "newbusiness"
            },
            {
              "propertyName": "hs_is_closed_won",
              "operator": "EQ",
              "value": "true"
            },
            {
              "propertyName": "closedate",
              "operator": "GT",
              "value": date
            }
          ]
        }
      ],
      "properties": ["amount", "hubspot_owner_id", "closedate", "pipeline", "hs_timestamp", "hs_is_closed_won", "dealtype"],
      "limit": 100
    })
  };
  var moreResults = true;
  var after = null;
  var row = 2;
  var owners = {};
  Logger.log('Started interacting with HubSpot.');
  Logger.log('Processing your request, please do not close the sheet...');
  while (moreResults) {
    if (after) {
      url = 'https://api.hubapi.com/crm/v3/objects/deals/search?after=' + after;
    }
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    var deals = data.results;
    for (var i = 0; i < deals.length; i++) {
      var deal = deals[i].properties;
      var timestamp = new Date(deal.closedate);
      var url1 = 'https://api.hubapi.com/owners/v2/owners/' + deal.hubspot_owner_id;
      var options = {
        'method': 'get',
        'headers': {
          'Authorization': 'Bearer ' + apiKey
        }
      };
      if(owners[deal.hubspot_owner_id]){
        owner = owners[deal.hubspot_owner_id]
      }
      else {
        var response1 = UrlFetchApp.fetch(url1, options);
        var data1 = JSON.parse(response1.getContentText());
        var owner = data1.firstName + " " + data1.lastName
        owners[deal.hubspot_owner_id] = owner;
      }
      sheet.getRange(row, 1).setValue(deal.amount)
      sheet.getRange(row, 2).setValue(owner);
      sheet.getRange(row, 3).setValue(deal.hubspot_owner_id ? 'Yes' : 'No');
      fdate = deal.closedate;
      finaldate = fdate.split('T');
      sheet.getRange(row, 4).setValue(finaldate[0]);
      ftime = finaldate[1].split('.');
      sheet.getRange(row, 5).setValue(ftime[0]);
      sheet.getRange(row, 6).setValue(deal.dealtype);
      sheet.getRange(row, 7).setValue(deal.pipeline);
      sheet.getRange(row, 8).setValue(deal.hs_is_closed_won);
      row++;
    }
    moreResults = !!data.paging && !!data.paging.next && !!data.paging.next.after;
    after = moreResults ? data.paging.next.after : null;
  }
}