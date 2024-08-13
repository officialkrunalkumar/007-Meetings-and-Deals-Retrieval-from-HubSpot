function fetchAndPopulateMeetings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.getRange(1, 1).setValue('Meeting Name');
  sheet.getRange(1, 2).setValue('Meeting Type');
  sheet.getRange(1, 3).setValue('Meeting Assigned To');
  sheet.getRange('1:1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  Logger.log('Headers of the sheet has been set and added!');
  var url = 'https://api.hubapi.com/crm/v3/objects/meetings?limit=100&properties=hs_meeting_title,hs_activity_type,hubspot_owner_id,hs_meeting_outcome,hs_timestamp';
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer Your_Authentication_Token'
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
    var meetings = data.results;
    for (var i = 0; i < meetings.length; i++) {
      var meeting = meetings[i].properties;
      var outcome = meeting.hs_meeting_outcome;
      var type = meeting.hs_activity_type
      var timestamp = new Date(meeting.hs_timestamp);
      var filterDate = new Date('2024-01-01');
      if (outcome === 'COMPLETED' && timestamp > filterDate && type != '' && type != null && type.includes('Zeni Overview')) {
        sheet.getRange(row, 1).setValue(meeting.hs_meeting_title);
        sheet.getRange(row, 2).setValue(meeting.hs_activity_type);
        var url1 = 'https://api.hubapi.com/owners/v2/owners/' + meeting.hubspot_owner_id;
        var options = {
          'method': 'get',
          'headers': {
            'Authorization': 'Bearer Your_Authentication_Token'
          }
        };
        var response1 = UrlFetchApp.fetch(url1, options);
        var data1 = JSON.parse(response1.getContentText());
        var owner = data1.firstName + " " + data1.lastName
        sheet.getRange(row, 3).setValue(owner);
        row++;
      }
    }
    if (data.paging && data.paging.next && data.paging.next.after) {
      after = data.paging.next.after;
    } else {
      moreResults = false;
    }
  }
  Logger.log("Meetings fetched and populated into Google Sheet.");
  sheet.getDataRange().setBorder(true, true, true, true, true, true);
  Logger.log('Borders have been added.');
  Logger.log('Execution finished. Your data is ready in the sheet. Thank you!');
}