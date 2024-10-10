function fetchAndPopulateMeetings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Meetings');
  var inputs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visual');
  var date = inputs.getRange('B3').getValue();
  Logger.log(date);
  sheet.clear();
  sheet.getRange(1, 1).setValue('Name');
  sheet.getRange(1, 2).setValue('Type');
  sheet.getRange(1, 3).setValue('Account Executive');
  sheet.getRange(1, 4).setValue('Start Date');
  sheet.getRange(1, 5).setValue('Start Time');
  sheet.getRange(1, 6).setValue('End Date');
  sheet.getRange(1, 7).setValue('End Time');
  sheet.getRange('1:1').setFontWeight('bold');
  sheet.setFrozenRows(1);
  Logger.log('Headers of the sheet have been set and added!');
  var apiKey = 'Your_Authentication_Token';
  var url = 'https://api.hubapi.com/crm/v3/objects/meetings/search';
  var filterDate = new Date(new Date().setDate(new Date().getDate() - 31));
  var endDate = new Date(new Date().setDate(new Date().getDate() - 1));
  var filterDateISO = filterDate.toISOString().split('T')[0] + "T00:00:00Z";
  var endDateISO = endDate.toISOString().split('T')[0] + "T23:59:59Z";
  Logger.log("Filter from: " + filterDateISO + " to: " + endDateISO);
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
              "propertyName": "hs_activity_type",
              "operator": "IN",
              "values": [
                'Zeni Overview - Inbound',
                'Zeni Overview - Outbound BDR',
                'Zeni Overview - Outbound AE',
                'Zeni Overview - Events AE',
                'Zeni Overview - Events BDR',
                'Zeni Overview - Events',
                'Zeni Overview - Events Booth',
                'Zeni Overview - VC Referral',
                'Zeni Overview - Customer Referral',
                'Zeni Overview - Employee Referral',
                'Zeni Overview - Inbound VC Referral',
                'Zeni Overview - Partnerships'
              ]
            },
            {
              "propertyName": "hs_meeting_outcome",
              "operator": "EQ",
              "value": "COMPLETED"
            },
            {
              "propertyName": "hs_meeting_start_time",
              "operator": "GTE",
              "value": filterDateISO
            },
            {
              "propertyName": "hs_meeting_start_time",
              "operator": "LTE",
              "value": endDateISO
            }
          ]
        }
      ],
      "properties": [
        "hs_meeting_title",
        "hs_activity_type",
        "hubspot_owner_id",
        "hs_meeting_outcome",
        "hs_timestamp",
        "hs_meeting_start_time",
        "hs_meeting_end_time"
      ],
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
      options.payload = JSON.stringify({
        "filterGroups": [
          {
            "filters": [
              {
                "propertyName": "hs_activity_type",
                "operator": "IN",
                "values": [
                  'Zeni Overview - Inbound',
                  'Zeni Overview - Outbound BDR',
                  'Zeni Overview - Outbound AE',
                  'Zeni Overview - Events AE',
                  'Zeni Overview - Events BDR',
                  'Zeni Overview - Events',
                  'Zeni Overview - Events Booth',
                  'Zeni Overview - VC Referral',
                  'Zeni Overview - Customer Referral',
                  'Zeni Overview - Employee Referral',
                  'Zeni Overview - Inbound VC Referral'
                ]
              },
              {
                "propertyName": "hs_meeting_outcome",
                "operator": "EQ",
                "value": "COMPLETED"
              },
              {
                "propertyName": "hs_meeting_start_time",
                "operator": "GTE",
                "value": filterDateISO
              },
              {
                "propertyName": "hs_meeting_start_time",
                "operator": "LTE",
                "value": endDateISO
              }
            ]
          }
        ],
        "properties": [
          "hs_meeting_title",
          "hs_activity_type",
          "hubspot_owner_id",
          "hs_meeting_outcome",
          "hs_timestamp",
          "hs_meeting_start_time",
          "hs_meeting_end_time"
        ],
        "limit": 100,
        "after": after
      });
    }
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    var meetings = data.results;
    for (var i = 0; i < meetings.length; i++) {
      var meeting = meetings[i].properties;
      sheet.getRange(row, 1).setValue(meeting.hs_meeting_title);
      sheet.getRange(row, 2).setValue(meeting.hs_activity_type);
      var url1 = 'https://api.hubapi.com/owners/v2/owners/' + meeting.hubspot_owner_id;
      var options1 = {
        'method': 'get',
        'headers': {
          'Authorization': 'Bearer ' + apiKey
        },
        'muteHttpExceptions': true
      };
      var owner;
      if (owners[meeting.hubspot_owner_id]) {
        owner = owners[meeting.hubspot_owner_id];
      } else {
        var response1 = UrlFetchApp.fetch(url1, options1);
        var data1 = JSON.parse(response1.getContentText());
        owner = data1.firstName + " " + data1.lastName;
        owners[meeting.hubspot_owner_id] = owner;
      }
      sheet.getRange(row, 3).setValue(owner);
      var sdate = meeting.hs_meeting_start_time;
      var finalSDate = sdate.split('T');
      sheet.getRange(row, 4).setValue(finalSDate[0]);
      var finalSTime = finalSDate[1].split(':');
      var finalUSTime = finalSTime[0] + ':' + finalSTime[1];
      sheet.getRange(row, 5).setValue(finalUSTime);
      var edate = meeting.hs_meeting_end_time;
      var finalEDate = edate.split('T');
      sheet.getRange(row, 6).setValue(finalEDate[0]);
      var finalETime = finalEDate[1].split(':');
      var finalUETime = finalETime[0] + ':' + finalETime[1];
      sheet.getRange(row, 7).setValue(finalUETime);
      row++;
    }
    moreResults = !!data.paging && !!data.paging.next && !!data.paging.next.after;
    after = moreResults ? data.paging.next.after : null;
  }
}