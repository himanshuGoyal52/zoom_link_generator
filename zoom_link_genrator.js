var ss = SpreadsheetApp.getActiveSpreadsheet();

var info_sheet = ss.getSheetByName("Information").getRange("C2:M2").getValues();
var email_idx = info_sheet[0][0];
var date_idx = info_sheet[0][1];
var topic_idx = info_sheet[0][2];
var start_time_idx = info_sheet[0][3];
var end_time_idx = info_sheet[0][4];
var mode_idx = info_sheet[0][5];
var passkey_idx = info_sheet[0][6];
var messages_idx = info_sheet[0][7];
var faculty_code_idx = info_sheet[0][8];
var telegram_group_idx = info_sheet[0][9];
var msg_sent_status_idx = info_sheet[0][10];

var int_to_alpha = {
  0: "A",
  1: "B",
  2: "C",
  3: "D",
  4: "E",
  5: "F",
  6: "G",
  7: "H",
  8: "I",
  9: "J",
  10: "K"
}

var Base64 = {
  _keyStr: "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",

  encode: function (input) {
    let output = "";
    let chr1, chr2, chr3, enc1, enc2, enc3, enc4;
    let i = 0;

    input = Base64._utf8_encode(input);

    while (i < input.length) {

      chr1 = input.charCodeAt(i++);
      chr2 = input.charCodeAt(i++);
      chr3 = input.charCodeAt(i++);

      enc1 = chr1 >> 2;
      enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
      enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
      enc4 = chr3 & 63;

      if (isNaN(chr2)) {
        enc3 = enc4 = 64;
      } else if (isNaN(chr3)) {
        enc4 = 64;
      }

      output = output +
        this._keyStr.charAt(enc1) + this._keyStr.charAt(enc2) +
        this._keyStr.charAt(enc3) + this._keyStr.charAt(enc4);
    }
    return output;
  },

  decode: function (input) {
    let output = "";
    let chr1, chr2, chr3;
    let enc1, enc2, enc3, enc4;
    let i = 0;

    input = input.replace(/[^A-Za-z0-9\+\/\=]/g, "");

    while (i < input.length) {

      enc1 = this._keyStr.indexOf(input.charAt(i++));
      enc2 = this._keyStr.indexOf(input.charAt(i++));
      enc3 = this._keyStr.indexOf(input.charAt(i++));
      enc4 = this._keyStr.indexOf(input.charAt(i++));

      chr1 = (enc1 << 2) | (enc2 >> 4);
      chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
      chr3 = ((enc3 & 3) << 6) | enc4;

      output = output + String.fromCharCode(chr1);

      if (enc3 != 64) {
        output = output + String.fromCharCode(chr2);
      }
      if (enc4 != 64) {
        output = output + String.fromCharCode(chr3);
      }
    }

    output = Base64._utf8_decode(output);

    return output;
  },

  _utf8_encode: function (string) {
    string = string.replace(/\r\n/g, "\n");
    let utftext = "";

    for (let n = 0; n < string.length; n++) {

      let c = string.charCodeAt(n);

      if (c < 128) {
        utftext += String.fromCharCode(c);
      }
      else if ((c > 127) && (c < 2048)) {
        utftext += String.fromCharCode((c >> 6) | 192);
        utftext += String.fromCharCode((c & 63) | 128);
      }
      else {
        utftext += String.fromCharCode((c >> 12) | 224);
        utftext += String.fromCharCode(((c >> 6) & 63) | 128);
        utftext += String.fromCharCode((c & 63) | 128);
      }
    }
    return utftext;
  },

  _utf8_decode: function (utftext) {
    let string = "";
    let i = 0;
    let c = c1 = c2 = 0;

    while (i < utftext.length) {

      c = utftext.charCodeAt(i);

      if (c < 128) {
        string += String.fromCharCode(c);
        i++;
      }
      else if ((c > 191) && (c < 224)) {
        c2 = utftext.charCodeAt(i + 1);
        string += String.fromCharCode(((c & 31) << 6) | (c2 & 63));
        i += 2;
      }
      else {
        c2 = utftext.charCodeAt(i + 1);
        c3 = utftext.charCodeAt(i + 2);
        string += String.fromCharCode(((c & 15) << 12) | ((c2 & 63) << 6) | (c3 & 63));
        i += 3;
      }
    }
    return string;
  }
}

function double_encode(val) {
  return Base64.encode(Base64.encode(val));
}

var telegram_maps_chatid = {};
var zoom_email_maps_credentials = {};
function onOpen() {
  let ui = SpreadsheetApp.getUi();

  ui.createMenu("Zoom")
    .addItem("Check Overlap", "check_lecture_overlap")
    .addItem("Create Zoom Link", "create_zoom_meet")
    .addItem("Create Zoom Link Bulk", "bulk_create_zoom_meet_v2")
    .addToUi();

  ui.createMenu("Telegram")
    .addItem("Send Message", "sendMessageOnTelegram")
    .addItem("Send Message Bulk", "bulk_sendMessageOnTelegram_v2")
    .addToUi();

  ui.createMenu("Doubts")
    .addItem("Single create" , "generate_doubts")
    .addItem("Bulk create" , "bulk_generate_doubts")
    .addToUi();


  let telegram_data = ss.getSheetByName("TelegramGroups").getDataRange().getDisplayValues();
  let zoom_data = ss.getSheetByName("Zoom Accounts").getDataRange().getDisplayValues();

  for (let i = 0; i < telegram_data.length; i++) {
    telegram_maps_chatid[telegram_data[i][0]] = telegram_data[i][1];
  }

  for (let i = 0; i < zoom_data.length; i++) {
    zoom_email_maps_credentials[zoom_data[i][1]] = [
      zoom_data[i][4],
      zoom_data[i][5],
      zoom_data[i][6]
    ];
  }

}

function open_a_dialog(data, dynamic_idx_overlap, message) {
  let message_html = HtmlService.createTemplateFromFile('Message');
  let ui = SpreadsheetApp.getUi();

  message_html.data = JSON.stringify(data);
  message_html.dynamic_idx_overlap = JSON.stringify(dynamic_idx_overlap);

  ui.showModalDialog(message_html.evaluate().setHeight(500).setWidth(900), message);
}

function tConvert(time) {
  time = time.toString().match(/^([01]\d|2[0-3])(:)([0-5]\d)(:[0-5]\d)?$/) || [time];

  if (time.length > 1) {
    time = time.slice(1);
    time[5] = +time[0] < 12 ? ' AM' : ' PM';
    time[0] = +time[0] % 12 || 12;
  }
  return time.join('');
}

function cal_duration(start, end) {
  let start_array = start.split(":");
  let end_array = end.split(":");

  let dur = 0;
  dur += (Number(end_array[0]) - Number(start_array[0])) * 60;
  dur += (Number(end_array[1]) - Number(start_array[1]));
  dur += (Number(end_array[2]) - Number(start_array[2])) / 60;

  return dur;
}

function get_seconds(time) {
  let time_arry = time.split(":");

  let sec = 0;
  sec += Number(time_arry[0]) * 3600;
  sec += Number(time_arry[1]) * 60;
  sec += Number(time_arry[2]);

  return sec;
}

function findOverlappingMeetings(meets) {
  let events = [];

  for (let meet of meets) {
    let [start, end, id] = meet;
    events.push({ time: start, type: 'start', id: id });
    events.push({ time: end, type: 'end', id: id });
  }

  events.sort((a, b) => a.time - b.time || (a.type === 'end' ? -1 : 1));

  let activeMeetings = new Set();
  let result = new Set();

  for (let event of events) {
    if (event.type === 'start') {
      activeMeetings.add(event.id);
      if (activeMeetings.size > 2) {
        for (let id of activeMeetings) {
          result.add(id);
        }
      }
    } else {
      activeMeetings.delete(event.id);
    }
  }

  return Array.from(result);
}

function cal_formated_duration(duration) {
  let hrs = Math.floor(duration / 60);
  let min = duration % 60;

  let ans = "";
  if (hrs) ans = hrs + " hrs ";

  if (min)
    ans += min + " min";


  return ans;
}

function cal_delay(told, real) {
  let told_arr = told.split(":")
  let real_arr = real.split(":");

  let told_min = Number(told_arr[0]) * 60 + Number(told_arr[1]);
  let real_min = Number(real_arr[0]) * 60 + Number(real_arr[1]);

  let ans = "";

  let delay_min = real_min - told_min;

  if (delay_min > 0) {
    ans = cal_formated_duration(delay_min);
  } else {
    ans = "No Delay";
  }

  return ans;

}

const swapElements = (array, index1, index2) => {
  let temp = array[index1];
  array[index1] = array[index2];
  array[index2] = temp;
}

function swapValuesTokey(j) {
  let res = {};
  for (let key in j) {
    res[j[key]] = key;
  }
  return res;
}

function ddmmyyyy_to_yyyymmdd(d) {
  let d_arr = d.split("-");

  if (d_arr.length === 3)
    return d_arr[2] + "-" + d_arr[1] + "-" + d_arr[0];
  else
    return;

}

function check_lecture_overlap(e = "c") {
  let ui = SpreadsheetApp.getUi();

  if (e == "uc") {
    ui.alert("🚧Under Construction!🚧 - This feature will be available soon.");
    return;
  } else if (SpreadsheetApp.getActiveSheet().getName() !== "Lectures") {
    ui.alert("Please use this option inside Lectures sheet");
    return;
  }

  let overlap_data = [];
  let is_overlap = false;
  let lectures = ss.getSheetByName("Lectures");
  let lectures_data_f = lectures.getDataRange().getDisplayValues();


  let lectures_data = [];

  let dynamic_idx_overlap = {
    "email_idx": email_idx,
    "date_idx": date_idx,
    "topic_idx": topic_idx,
    "start_time_idx": start_time_idx,
    "end_time_idx": end_time_idx,
    "mode_idx": mode_idx,
    "passkey_idx": passkey_idx,
    "messages_idx": messages_idx,
    "faculty_code_idx": faculty_code_idx,
    "telegram_group_idx": telegram_group_idx,
    "msg_sent_status_idx": msg_sent_status_idx
  }
  let dynamic_idx_overlap_rev = swapValuesTokey(dynamic_idx_overlap);

  for (let i = 0; i < lectures_data_f.length; i++) {

    if (lectures_data_f[i][dynamic_idx_overlap.email_idx] != "#N/A" && lectures_data_f[i][dynamic_idx_overlap.email_idx].toString().trim().length != 0) {
      swapElements(lectures_data_f[i], dynamic_idx_overlap.date_idx, 1); // 0 , 1
      swapElements(lectures_data_f[i], 0, dynamic_idx_overlap.email_idx); // 0 , 4
      let arru = lectures_data_f[i].slice(0, 11);
      arru.push(i + 1);
      lectures_data.push(arru);
    }
  }

  /* swaping the index */
  let tempo = dynamic_idx_overlap.date_idx;
  dynamic_idx_overlap.date_idx = 1;
  dynamic_idx_overlap[dynamic_idx_overlap_rev[1]] = tempo;

  dynamic_idx_overlap_rev = swapValuesTokey(dynamic_idx_overlap);

  let tempo2 = dynamic_idx_overlap.email_idx;
  dynamic_idx_overlap.email_idx = 0;
  dynamic_idx_overlap[dynamic_idx_overlap_rev[0]] = tempo2;

  dynamic_idx_overlap_rev = swapValuesTokey(dynamic_idx_overlap);
  /**************************/


  let lectures_data_c = [];
  for (let i = 0; i < lectures_data.length; i++) {
    let rrow = [];
    for (let j = 0; j < lectures_data[i].length; j++) {
      rrow.push(lectures_data[i][j]);
    }
    rrow.push(i + 1);
    lectures_data_c.push(rrow);
  }

  for (let i = 0; i < lectures_data.length; i++) lectures_data[i].push(i + 1);
  lectures_data.shift();
  let sorted_data = lectures_data.sort();


  let i;
  for (i = 0; i < sorted_data.length;) {
    let j = i;
    while (j < sorted_data.length && sorted_data[j][dynamic_idx_overlap.email_idx] == sorted_data[i][dynamic_idx_overlap.email_idx] && sorted_data[j][dynamic_idx_overlap.date_idx] == sorted_data[i][dynamic_idx_overlap.date_idx]) {
      j++;
    }

    if (j - i > 2) {
      let time_spans = [];
      for (let x = i; x < j; x++) {
        time_spans.push([get_seconds(sorted_data[x][dynamic_idx_overlap.start_time_idx]), get_seconds(sorted_data[x][dynamic_idx_overlap.end_time_idx]), sorted_data[x][sorted_data[x].length - 1]]);
      }

      let sorted_time_spans = time_spans.sort();
      let overlap_row = findOverlappingMeetings(sorted_time_spans);

      if (overlap_row.length > 0) {
        overlap_data.push(overlap_row);
        is_overlap = true;
      }

    }

    i = j;
  }

  let real_overlap = [];

  for (let i = 0; i < overlap_data.length; i++) {
    let one_row = [];
    for (let j = 0; j < overlap_data[i].length; j++) {
      one_row.push(lectures_data_c[overlap_data[i][j] - 1]);
    }
    real_overlap.push(one_row);
  }

  if (is_overlap) {
    open_a_dialog(real_overlap, dynamic_idx_overlap, "Below listed meeting are overlaping!! please change the emails");
  } else {
    open_a_dialog([], {}, "There are no overlaps!! you can generate zoom links")
  }
}

function get_zoom_access_token(account_id, client_id, client_secret) {

  let access_token_res = null;
  try {
    let details = {
      'grant_type': "account_credentials",
      'account_id': account_id,
    };

    let formBody = [];
    for (let property in details) {
      let encodedKey = encodeURIComponent(property);
      let encodedValue = encodeURIComponent(details[property]);
      formBody.push(encodedKey + "=" + encodedValue);
    }
    formBody = formBody.join("&");

    access_token_res = JSON.parse(UrlFetchApp.fetch('https://zoom.us/oauth/token', {
      "method": "POST",
      "headers": {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': `Basic ${Base64.encode(client_id + ":" + client_secret)}`
      },
      "payload": formBody
    })
    )

  } catch (e) {
    return null;
  }

  return access_token_res;
}

function create_zoom_meet(type = "single", row = null) {

  let ui = SpreadsheetApp.getUi();
  if (type === "single" && SpreadsheetApp.getActiveSheet().getName() !== "Lectures") {
    ui.alert("Please use this option inside Lectures sheet");
    return;
  }

  let zoom_accounts_data = ss.getSheetByName("Zoom Accounts").getDataRange().getDisplayValues();
  let msg_template_data = ss.getSheetByName("Message").getDataRange().getValues();
  let lectures = ss.getSheetByName("Lectures");

  let cell = lectures.getActiveCell();
  if (row === null && type === "single") {
    row = cell.getRowIndex();
  }
  if (type === "single" && (cell.getValue().toString().trim().length == 0 || row == 1)) {
    ui.alert("Please select a row with meeting data");
    return;
  }

  let meeting_data = lectures.getRange(row, 1, 1, lectures.getMaxColumns()).getDisplayValues();
  let client_id = "", client_secret = "", account_id = "";


  let msg_template_found = true;
  if (meeting_data[0][mode_idx].toString().trim().length == 0) {
    msg_template_found = false;
  }

  let found = false;

  if (meeting_data[0][email_idx] in zoom_email_maps_credentials) {
    found = true;
    client_id = zoom_email_maps_credentials[meeting_data[0][email_idx]][0]
    client_secret = zoom_email_maps_credentials[meeting_data[0][email_idx]][1]
    account_id = zoom_email_maps_credentials[meeting_data[0][email_idx]][2]
  } else {
    for (let i = 0; i < zoom_accounts_data.length; i++) {
      if (zoom_accounts_data[i][1] == meeting_data[0][email_idx]) {
        found = true;
        client_id = zoom_accounts_data[i][4];
        client_secret = zoom_accounts_data[i][5];
        account_id = zoom_accounts_data[i][6];
        break;
      }
    }
  }


  if (found === false) {
    if (type === "single")
      ui.alert("There is no such email id in \"Zoom Accounts\" sheet");
    return "There is no such email id in \"Zoom Accounts\" sheet";
  }
  if (client_id.length == 0 || client_secret.length == 0 || account_id.length == 0) {
    if (type === "single")
      ui.alert("Please add full information of this account in \"Zoom Accounts\" sheet");
    return "Please add full information of this account in \"Zoom Accounts\" sheet";
  }

  let access_token_res = null;
  try {
    let details = {
      'grant_type': "account_credentials",
      'account_id': account_id,
    };

    let formBody = [];
    for (let property in details) {
      let encodedKey = encodeURIComponent(property);
      let encodedValue = encodeURIComponent(details[property]);
      formBody.push(encodedKey + "=" + encodedValue);
    }
    formBody = formBody.join("&");

    access_token_res = JSON.parse(UrlFetchApp.fetch('https://zoom.us/oauth/token', {
      "method": "POST",
      "headers": {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': `Basic ${Base64.encode(client_id + ":" + client_secret)}`
      },
      "payload": formBody
    })
    )

  } catch (e) {
    if (type === "single")
      ui.alert(e.message);
    return e.message;
  }


  let created_meeting_data;

  if (access_token_res) {
    try {
      let body = {
        "agenda": meeting_data[0][topic_idx],
        "default_password": false,
        "duration": cal_duration(meeting_data[0][start_time_idx], meeting_data[0][end_time_idx]),
        "password": meeting_data[0][passkey_idx],
        "pre_schedule": false,
        "schedule_for": meeting_data[0][email_idx],
        "settings": {
          "allow_multiple_devices": false,
          "approval_type": 2,
          "audio": "voip",
          "authentication_domains": "bakliwaltutorials.com,bakliwaltutorialsiit.onmicrosoft.com",
          "auto_recording": "cloud",
          "contact_email": meeting_data[0][email_idx],
          "contact_name": "BT-Lecture",
          "email_notification": true,
          "encryption_type": "enhanced_encryption",
          "focus_mode": true,
          "host_video": false,
          "jbh_time": 0,
          "join_before_host": false,
          "meeting_authentication": true,
          "mute_upon_entry": true,
          "participant_video": false,
          "use_pmi": false,
          "waiting_room": false,
          "watermark": false,
          "internal_meeting": false,
          "continuous_meeting_chat": {
            "enable": true,
            "auto_add_invited_external_users": false
          },
          "participant_focused_meeting": false,
          "auto_start_meeting_summary": false,
          "auto_start_ai_companion_questions": false
        },
        "start_time": `${ddmmyyyy_to_yyyymmdd(meeting_data[0][date_idx])}T${meeting_data[0][start_time_idx]}`,
        "timezone": "Asia/Calcutta",
        "topic": meeting_data[0][topic_idx],
        "type": 2
      }

      let created_meeting_data_res = UrlFetchApp.fetch(`https://api.zoom.us/v2/users/${meeting_data[0][email_idx]}/meetings`, {
        "method": "POST",
        "headers": {
          'Content-Type': 'application/json',
          'Authorization': `${access_token_res.token_type} ${access_token_res.access_token}`
        },
        "payload": JSON.stringify(body),
      });

      if (created_meeting_data_res.getResponseCode() === 429) {
        if (type === "single")
          ui.alert("Too Many Requests");
        return "Too Many Requests";
      }
      if (created_meeting_data_res.getResponseCode() === 400) {
        if (type === "single")
          ui.alert("Bad Request");
        return "Bad Request";
      }
      if (created_meeting_data_res.getResponseCode() === 404) {
        if (type === "single")
          ui.alert("Not Found, (User does not exist: {userId}.)");
        return "Not Found, (User does not exist: {userId}.)";
      }

      created_meeting_data = JSON.parse(created_meeting_data_res);

      let range = lectures.getRange(`${int_to_alpha[messages_idx]}${row}`);
      let meet_credentials = created_meeting_data.join_url + '\nMeeting ID: ' + created_meeting_data.id + '\nPasscode: ' + created_meeting_data.password;



      if (msg_template_found) {
        let final_msg = "";
        for (let i = 1; i < msg_template_data.length; i++) {
          if (msg_template_data[i][0] == meeting_data[0][mode_idx]) {
            final_msg = msg_template_data[i][1];
            break;
          }
        }

        final_msg = final_msg.replace("[DATE]", new Date(ddmmyyyy_to_yyyymmdd(meeting_data[0][date_idx])).toDateString());
        final_msg = final_msg.replace("[CODE]", meeting_data[0][faculty_code_idx]);
        final_msg = final_msg.replace("[START_TIME]", tConvert(meeting_data[0][start_time_idx]));
        final_msg = final_msg.replace("[END_TIME]", tConvert(meeting_data[0][end_time_idx]));
        final_msg += "\n\n" + meet_credentials + "\n\n" + "<b>BT Team</b>";

        range.setValue(final_msg);



      } else {
        range.setValue(meet_credentials);
      }


      let reports_helping_meta_data_sheet = ss.getSheetByName("ReportsHelpingMetaData");
      let rhmd_last_row = reports_helping_meta_data_sheet.getLastRow() + 1;
      reports_helping_meta_data_sheet.getRange(`A${rhmd_last_row}:H${rhmd_last_row}`).setValues([
        [meeting_data[0][date_idx],
        meeting_data[0][faculty_code_idx],
        meeting_data[0][email_idx],
        created_meeting_data.id,
        meeting_data[0][start_time_idx],
        meeting_data[0][end_time_idx],
        meeting_data[0][topic_idx],
          ""
        ]
      ]);


    } catch (e) {
      if (type === "single")
        ui.alert(e.message);
      return e.message;
    }
  }

  return "DONE";

}

function bulk_create_zoom_meet_v2(e = "c") {
  let ui = SpreadsheetApp.getUi();

  if (e === "uc") {
    ui.alert("🚧Under Construction!🚧 - This feature will be available soon.");
    return;
  }

  if (SpreadsheetApp.getActiveSheet().getName() !== "Lectures") {
    ui.alert("Please use this option inside Lectures sheet");
    return;
  }

  let input = ui.prompt("Please enter the range (example: 3-30)");

  let button = input.getSelectedButton();
  let lecture_sheet = ss.getSheetByName("Lectures");

  if (button === ui.Button.OK) {
    let range = input.getResponseText();
    let range_arr = range.split("-");
    if (range_arr.length !== 2) {
      ui.alert("Enter the valid range");
      return;
    }

    if (isNaN(range_arr[0].toString().trim()) || isNaN(range_arr[1].toString().trim())) {
      ui.alert("Enter the valid range");
      return;
    }

    let start = Number(range_arr[0]);
    let end = Number(range_arr[1]);

    if (end < start) {
      ui.alert("Start should be lesser than end");
      return;
    }

    if (start === 1 || end === 1 || start === 2 || end === 2) {
      ui.alert("please do not enter value 1 or 2 (i.e. Header)");
      return;
    }

    if (end - start + 1 > 50) {
      ui.alert("The range should be smaller than 50.");
      return;
    }

    let all_done = true;

    for (let row = start; row <= end; row++) {
      if (lecture_sheet.getRange(`${int_to_alpha[email_idx]}${row}`).getValue().toString().trim().length > 0) {
        let res = create_zoom_meet("bulk", row);
        if (res !== "DONE") {
          all_done = false;
          let range = lecture_sheet.getRange(`${int_to_alpha[messages_idx]}${row}`);
          range.setValue("error : " + res);
        }
      }
    }

    if (!all_done) {
      ui.alert("Some meeting are not created kindly check.");
    }



  } else if (button === ui.Button.CLOSE) {
    Logger.log("The user clicked the [X] button and closed the prompt dialog.");
  }


}

function bulk_create_zoom_meet(e = "uc") {
  let ui = SpreadsheetApp.getUi();
  if (e === "uc") {
    ui.alert("❗Attention❗ - This feature is shifted to the dropdown of Zoom.");
    return;
  }

  let lecture_sheet = ss.getSheetByName("Lectures");
  let start = lecture_sheet.getRange("L2").getValue();
  let end = lecture_sheet.getRange("M2").getValue();

  if (start.toString().trim().length == 0 || end.toString().trim().length == 0) {
    ui.alert("Please enter the start and end values of Zoom Range");
    return;
  }

  if (end < start) {
    ui.alert("Start should be lesser than end");
    return;
  }

  if (start == 1 || end == 1) {
    ui.alert("please do not enter value 1 (i.e. Header)");
    return;
  }

  if (end - start + 1 > 50) {
    ui.alert("The range should be smaller than 50.");
    return;
  }

  let all_done = true;

  for (let row = start; row <= end; row++) {
    if (lecture_sheet.getRange(`${int_to_alpha[email_idx]}${row}`).getValue().toString().trim().length > 0) {
      let res = create_zoom_meet("bulk", row);
      if (res !== "DONE") {
        all_done = false;
        let range = lecture_sheet.getRange(`${int_to_alpha[messages_idx]}${row}`);
        range.setValue("error : " + res);
      }
    }
  }

  if (!all_done) {
    ui.alert("Some meeting are not created kindly check.");
  }

}

function sendMessageOnTelegram(type = "single", row = null) {
  let ui = SpreadsheetApp.getUi();
  if (type === "single" && SpreadsheetApp.getActiveSheet().getName() !== "Lectures") {
    ui.alert("Please use this option inside Lectures sheet");
    return;
  }

  let chat_id = null;

  let lectures = ss.getSheetByName("Lectures");
  let telegramGrps = ss.getSheetByName("TelegramGroups").getDataRange().getValues();

  let cell = lectures.getActiveCell();

  if (type === "single" && (cell.getValue().toString().trim().length == 0 || row == 1)) {
    ui.alert("Please select a row with meeting data");
    return;
  }

  if (row === null)
    row = cell.getRowIndex();

  let meeting_data = lectures.getRange(row, 1, 1, lectures.getMaxColumns()).getDisplayValues();

  if (meeting_data[0][telegram_group_idx].toString().trim().length == 0) {
    if (type === "single")
      ui.alert("No telegram group is selected");
    return "No telegram group is selected";
  }

  /***************************************************************************************/
  if (meeting_data[0][telegram_group_idx] in telegram_maps_chatid) {
    chat_id = telegram_maps_chatid[meeting_data[0][telegram_group_idx]];
  } else {
    for (let i = 0; i < telegramGrps.length; i++) {
      if (telegramGrps[i][0] == meeting_data[0][telegram_group_idx]) {
        chat_id = telegramGrps[i][1];
        break;
      }
    }
  }

  if (chat_id === null) {
    if (type === "single")
      ui.alert("There is no such telegram group exits \"TelegramGroups\" sheet");
    return "There is no such telegram group exits \"TelegramGroups\" sheet";
  }

  if (chat_id.toString().trim().length === 0 || isNaN(chat_id)) {
    if (type === "single")
      ui.alert("Please add Chat Id of this telegram group in \"TelegramGroups\" sheet.");
    return "Please add Chat Id of this telegram group in \"TelegramGroups\" sheet.";
  }
  /***************************************************************************************/

  let token = ss.getSheetByName("Information").getRange("A6").getValue();
  let text_message = lectures.getRange(`${int_to_alpha[messages_idx]}${row}`).getValue();
  if (text_message.toString().trim().length === 0) {
    if (type === "single")
      ui.alert("Please genrate zoom meet or message first");
    return "Please genrate zoom meet or message first";
  } else if (text_message.toString().substring(0, 5) == "error") {
    if (type === "single")
      ui.alert("Please genrate zoom meet or message first");
    return "Please genrate zoom meet or message first";
  } else if (lectures.getRange(`${int_to_alpha[mode_idx]}${row}`).getValue().toString().trim().length === 0) {
    if (type === "single")
      ui.alert("No message mode is selected");
    return "No message mode is selected";
  }

  text_message = encodeURI(text_message);

  try {
    let url = `https://api.telegram.org/bot${token}/sendMessage?chat_id=${chat_id}&text=${text_message}&parse_mode=HTML`;
    let res_msg = JSON.parse(UrlFetchApp.fetch(url));
    if (res_msg.ok == true) {
      if (type === "single") {
        ui.alert("message sent!!");
        lectures.getRange(`${int_to_alpha[msg_sent_status_idx]}${row}`).setValue("message sent!!");
      }
      return "message sent!!";
    } else {
      if (res_msg.description) {
        if (type === "single")
          ui.alert(res_msg.description);
        return res_msg.description;
      }
      else {
        if (type === "single")
          ui.alert("Some error occured");
        return "Some error occured";
      }
    }

  } catch (e) {
    if (type === "single")
      ui.alert(e.message);
    return e.message;
  }
}

function bulk_sendMessageOnTelegram_v2(e = "c") {
  let ui = SpreadsheetApp.getUi();

  if (e === "uc") {
    ui.alert("🚧Under Construction!🚧 - This feature will be available soon.");
    return;
  }

  if (SpreadsheetApp.getActiveSheet().getName() !== "Lectures") {
    ui.alert("Please use this option inside Lectures sheet");
    return;
  }

  let input = ui.prompt("Please enter the range (example: 3-30)");

  let button = input.getSelectedButton();
  let lecture_sheet = ss.getSheetByName("Lectures");

  if (button === ui.Button.OK) {
    let range = input.getResponseText();
    let range_arr = range.split("-");
    if (range_arr.length !== 2) {
      ui.alert("Enter the valid range");
      return;
    }

    if (isNaN(range_arr[0].toString().trim()) || isNaN(range_arr[1].toString().trim())) {
      ui.alert("Enter the valid range");
      return;
    }

    let start = Number(range_arr[0]);
    let end = Number(range_arr[1]);

    if (start === 1 || end === 1 || start === 2 || end === 2) {
      ui.alert("please do not enter value 1 or 2 (i.e. Header)");
      return;
    }

    if (end < start) {
      ui.alert("Start should be lesser than end");
      return;
    }

    if (end - start + 1 > 50) {
      ui.alert("The range should be smaller than 50.");
      return;
    }

    let all_done = true;
    for (let row = start; row <= end; row++) {
      if (lecture_sheet.getRange(`${int_to_alpha[email_idx]}${row}`).getValue().toString().trim().length > 0) {
        let res = sendMessageOnTelegram("bulk", row);
        lecture_sheet.getRange(`${int_to_alpha[msg_sent_status_idx]}${row}`).setValue(res);
        if (res !== "message sent!!") {
          all_done = false;
        }
      }
    }

    if (!all_done) {
      ui.alert("Some messages are not sent kindly check.");
      return;
    }


  } else if (button === ui.Button.CLOSE) {
    Logger.log("The user clicked the [X] button and closed the prompt dialog.");
  }
}

function bulk_sendMessageOnTelegram(e = "uc") {
  let ui = SpreadsheetApp.getUi();
  if (e === "uc") {
    ui.alert("❗Attention❗ - This feature is shifted to the dropdown of Telegram.");
    return;
  }

  let lecture_sheet = ss.getSheetByName("Lectures");
  let start = lecture_sheet.getRange("L3").getValue();
  let end = lecture_sheet.getRange("M3").getValue();

  if (start.toString().trim().length == 0 || end.toString().trim().length == 0) {
    ui.alert("Please enter the start and end values of Telegram Range");
    return;
  }

  if (start == 1 || end == 1) {
    ui.alert("please do not enter value 1 (i.e. Header)");
    return;
  }

  if (end < start) {
    ui.alert("Start should be lesser than end");
    return;
  }

  if (end - start + 1 > 50) {
    ui.alert("The range should be smaller than 50.");
    return;
  }

  let all_done = true;
  for (let row = start; row <= end; row++) {
    if (lecture_sheet.getRange(`${int_to_alpha[email_idx]}${row}`).getValue().toString().trim().length > 0) {
      let res = sendMessageOnTelegram("bulk", row);
      lecture_sheet.getRange(`${int_to_alpha[msg_sent_status_idx]}${row}`).setValue(res);
      if (res !== "message sent!!") {
        all_done = false;
      }
    }
  }

  if (!all_done) {
    ui.alert("Some messages are not sent kindly check.");
    return;
  }

}

function genrate_chat_id(e = "c") {
  let ui = SpreadsheetApp.getUi();
  if (e === "uc") {
    ui.alert("🚧Under Construction!🚧 - This feature will be available soon.");
    return;
  }


  let token = ss.getSheetByName("Information").getRange("A6").getValue();

  let updates_res;
  let process = {};

  try {

    let url = `https://api.telegram.org/bot${token}/getUpdates`;

    let last_update_id = ss.getSheetByName("Information").getRange("A7").getValue();
    if (last_update_id) {
      url += `?offset=${last_update_id}`;
    }

    updates_res = JSON.parse(UrlFetchApp.fetch(url));
    let results = updates_res.result;


    if (results.length === 100) {
      ss.getSheetByName("Information").getRange("A7").setValue(results[99].update_id + 1);
    } else {
      ss.getSheetByName("Information").getRange("A7").setValue("");
    }

    for (let i = 0; i < results.length; i++) {
      if ("my_chat_member" in results[i]) {
        if (results[i].my_chat_member.chat.title && results[i].my_chat_member.chat.id)
          process[results[i].my_chat_member.chat.title] = Number(results[i].my_chat_member.chat.id);
      }
    }


  } catch (e) {
    ui.alert(e.message);
    return;
  }


  let telegrpsheet = ss.getSheetByName("TelegramGroups");
  let telegrps = telegrpsheet.getDataRange().getValues();

  let error = false;

  for (let i = 1; i < telegrps.length; i++) {
    if (telegrps[i][1].toString().trim().length > 0 && !isNaN(telegrps[i][1])) continue;

    if (process[telegrps[i][0]]) {
      telegrpsheet.getRange(`B${i + 1}`).setValue(process[telegrps[i][0]]);
    } else {
      error = true;
      telegrpsheet.getRange(`B${i + 1}`).setValue("Please add @bakliwal_msg_bot to the group first");
    }
  }

  if (error) {
    ui.alert("Chat Id's of some of the groups is not genrated. please check");
  }

}

function report_generation(e = "c") {
  let ui = SpreadsheetApp.getUi();
  if (e === "uc") {
    ui.alert("🚧Under Construction!🚧 - This feature will be available soon.");
    return;
  }

  let past_meetings_reports_sheet = ss.getSheetByName("Past Meetings Reports");

  let rhmd = ss.getSheetByName("ReportsHelpingMetaData").getDataRange().getDisplayValues();
  let zoom_accounts_data = ss.getSheetByName("Zoom Accounts").getDataRange().getDisplayValues();

  let zoom_cred_obj = {};
  for (let j = 0; j < zoom_accounts_data.length; j++) {
    zoom_cred_obj[zoom_accounts_data[j][1]] = [zoom_accounts_data[j][4], zoom_accounts_data[j][5], zoom_accounts_data[j][6]];
  }

  // Date ,	Facutly Code ,	Email ,	Meeting ID ,	Start Time ,	End Time , Meeting Topic ,	Report Done?

  let all_generated = true;

  for (let i = 1; i < rhmd.length; i++) {
    if (rhmd[i][7].toString().trim().length === 0) {
      all_generated = false;
      let access_token_res = get_zoom_access_token(zoom_cred_obj[rhmd[i][2]][2], zoom_cred_obj[rhmd[i][2]][0], zoom_cred_obj[rhmd[i][2]][1]);

      if (access_token_res) {
        try {
          let meeting_id = rhmd[i][3]
          if (rhmd[i][3].toString().includes("/")) meeting_id = double_encode(rhmd[i][3]);

          let meeting_info_data_res = UrlFetchApp.fetch(`https://api.zoom.us/v2/past_meetings/${meeting_id}`, {
            "method": "GET",
            "headers": {
              'Content-Type': 'application/json',
              'Authorization': `${access_token_res.token_type} ${access_token_res.access_token}`
            }
          });

          let meeting_info_data = JSON.parse(meeting_info_data_res);

          // Logger.log(meeting_info_data);

          if (meeting_info_data_res.getResponseCode() === 200) {

            let last_row = past_meetings_reports_sheet.getLastRow() + 1;

            try {
              past_meetings_reports_sheet.getRange(`A${last_row}:L${last_row}`).setValues([[
                rhmd[i][0],
                rhmd[i][1],
                rhmd[i][3],
                rhmd[i][6],
                rhmd[i][4],
                rhmd[i][5],
                new Date(meeting_info_data.start_time).toTimeString().split(" ")[0],
                new Date(meeting_info_data.end_time).toTimeString().split(" ")[0],
                cal_formated_duration(cal_duration(rhmd[i][4], rhmd[i][5])),
                cal_formated_duration(meeting_info_data.duration),
                cal_delay(rhmd[i][4], new Date(meeting_info_data.start_time).toTimeString().split(" ")[0]),
                meeting_info_data.participants_count
              ]]);

            } catch (e) {
              Logger.log(e.message);
            }


            ss.getSheetByName("ReportsHelpingMetaData").getRange("H" + (i + 1)).setValue("Yes");
            ss.getSheetByName("ReportsHelpingMetaData").getRange("I" + (i + 1)).setValue("");
          } else {
            ss.getSheetByName("ReportsHelpingMetaData").getRange("I" + (i + 1)).setValue(meeting_info_data.message);
          }
        } catch (e) {
          ss.getSheetByName("ReportsHelpingMetaData").getRange("I" + (i + 1)).setValue(e.message);
        }
      }


    }
  }

  if (all_generated) {
    ui.alert("All reports are generated!");
  }

}

function bulk_sendMessage(e = "c") {
  let ui = SpreadsheetApp.getUi();
  if (e === "uc") {
    ui.alert("🚧Under Construction!🚧 - This feature will be available soon.");
    return;
  }

  let telegram_sheet = ss.getSheetByName("Telegram Message Bulk");
  let t_s_data = telegram_sheet.getDataRange().getValues();
  let telegram_groups = ss.getSheetByName("TelegramGroups").getDataRange().getValues();
  let token = ss.getSheetByName("Information").getRange("A6").getValue();

  if (telegram_sheet.getRange("C2").getValue().toString().trim().length === 0) {
    ui.alert("There is no message to send please enter the required message in C2 cell.");
    return;
  }

  let text_message = encodeURI(telegram_sheet.getRange("C2").getValue());

  let telegram_obj = {};
  for (let i = 1; i < telegram_groups.length; i++) {
    if (!isNaN(telegram_groups[i][1]))
      telegram_obj[telegram_groups[i][0]] = telegram_groups[i][1];
  }

  for (let i = 1; i < t_s_data.length; i++) {
    if (t_s_data[i][0] in telegram_obj) {
      let chat_id = telegram_obj[t_s_data[i][0]];
      try {
        let url = `https://api.telegram.org/bot${token}/sendMessage?chat_id=${chat_id}&text=${text_message}&parse_mode=HTML`;
        let res_msg = JSON.parse(UrlFetchApp.fetch(url));
        if (res_msg.ok == true) {
          telegram_sheet.getRange(`B${i + 1}`).setValue("message sent!!");
        } else {
          if (res_msg.description) {
            telegram_sheet.getRange(`B${i + 1}`).setValue(res_msg.description);
          }
          else {
            telegram_sheet.getRange(`B${i + 1}`).setValue("Some error occured");
          }
        }

      } catch (e) {
        telegram_sheet.getRange(`B${i + 1}`).setValue(e.message);
      }
    } else {
      telegram_sheet.getRange(`B${i + 1}`).setValue("Chat ID does not exist for this group.");
    }
  }

}

function onEdit(e) {
 let ss = SpreadsheetApp.getActiveSpreadsheet();
 let activeCell = ss.getActiveCell();

 // {authMode=LIMITED, user=bakliwaltutorialswebapps@gmail.com, value=C26 Lloyds Noon COC, range=Range, source=Spreadsheet, oldValue=FACC 25Completely Online - COC [Wed + Sat]}
  if((activeCell.getColumn() == 6 || activeCell.getColumn() == 5) && ss.getActiveSheet().getName() == "Doubts") {
    let newValue = e.value;
    let oldValue = e.oldValue;

    if (!newValue) {
      activeCell.setValue("");
    } else {
      if (!oldValue) {
        activeCell.setValue(newValue);
      } else {
        activeCell.setValue(oldValue + "|~|" + newValue);
      }
    }
  }

  if(activeCell.getColumn() == 7 && ss.getActiveSheet().getName() == "Lectures") {
    let newValue = e.value;
    let oldValue = e.oldValue;

    if (!newValue) {
      activeCell.setValue("");
    } else {
      if (!oldValue) {
        activeCell.setValue(newValue);
      } else {
        activeCell.setValue(oldValue + "|~|" + newValue);
      }
    }
  }

}

function sendMessageOnTelegram_v2(type = "single", row = null) {
  let ui = SpreadsheetApp.getUi();
  if (type === "single" && SpreadsheetApp.getActiveSheet().getName() !== "Lectures") {
    ui.alert("Please use this option inside Lectures sheet");
    return;
  }  

  let lectures = ss.getSheetByName("Lectures");
  let telegramGrps = ss.getSheetByName("TelegramGroups").getDataRange().getValues();

  let cell = lectures.getActiveCell();

  if (type === "single" && (cell.getValue().toString().trim().length == 0 || row == 1)) {
    ui.alert("Please select a row with meeting data");
    return;
  }

  if (row === null)
    row = cell.getRowIndex();

  let meeting_data = lectures.getRange(row, 1, 1, lectures.getMaxColumns()).getDisplayValues();

  if (meeting_data[0][telegram_group_idx].toString().trim().length == 0) {
    if (type === "single")
      ui.alert("No telegram group is selected");
    return "No telegram group is selected";
  }

  let token = ss.getSheetByName("Information").getRange("A6").getValue();
  let text_message = lectures.getRange(`${int_to_alpha[messages_idx]}${row}`).getValue();
  if (text_message.toString().trim().length === 0) {
    if (type === "single")
      ui.alert("Please genrate zoom meet or message first");
    return "Please genrate zoom meet or message first";
  } else if (text_message.toString().substring(0, 5) == "error") {
    if (type === "single")
      ui.alert("Please genrate zoom meet or message first");
    return "Please genrate zoom meet or message first";
  } else if (lectures.getRange(`${int_to_alpha[mode_idx]}${row}`).getValue().toString().trim().length === 0) {
    if (type === "single")
      ui.alert("No message mode is selected");
    return "No message mode is selected";
  }

  

  let telegram_groups_arr = meeting_data[0][telegram_group_idx].toString().split("|~|");

  let final_msg = "";

  for(let j=0 ; j<telegram_groups_arr.length ; j++){
    let chat_id = null;

      if (telegram_groups_arr[j] in telegram_maps_chatid) {
        chat_id = telegram_maps_chatid[telegram_groups_arr[j]];
      } else {
        for (let i = 0; i < telegramGrps.length; i++) {
          if (telegramGrps[i][0] == telegram_groups_arr[j]) {
            chat_id = telegramGrps[i][1];
            break;
          }
        }
      }

      if (chat_id === null) {
        final_msg += telegram_groups_arr[j] + " : " + "There is no such telegram group exits \"TelegramGroups\" sheet\n";
        continue;
      }

      if (chat_id.toString().trim().length === 0 || isNaN(chat_id)) {
        final_msg += telegram_groups_arr[j] + " : " + "Please add Chat Id of this telegram group in \"TelegramGroups\" sheet.\n";
        continue;
      }

      text_message = encodeURI(text_message);

      try {
        let url = `https://api.telegram.org/bot${token}/sendMessage?chat_id=${chat_id}&text=${text_message}&parse_mode=HTML`;
        let res_msg = JSON.parse(UrlFetchApp.fetch(url));
        if (res_msg.ok == true) {
          final_msg += telegram_groups_arr[j] + " : " + "message sent!!\n";
          continue;

        } else {
          if (res_msg.description) {
            final_msg += telegram_groups_arr[j] + " : " + res_msg.description + "\n";
            continue;
          }
          else {
            final_msg += telegram_groups_arr[j] + " : " + "Some error occured\n";
            continue;
          }
        }
      } catch (e) {
        final_msg += telegram_groups_arr[j] + " : " + e.message + "\n";
        continue;
      }
  }

  if (type === "single") {
    ui.alert(final_msg);
    lectures.getRange(`${int_to_alpha[msg_sent_status_idx]}${row}`).setValue(final_msg);
    return;
  }

  return final_msg;
}

function clean_up_trigger(sheet_name , date_idx){
  let sheet = ss.getSheetByName(sheet_name);
  let ini_len = sheet.getLastRow();
  let last_col_char = String.fromCharCode(65 + (sheet.getLastColumn()-1));
  let data = sheet.getRange(`A2:${last_col_char}${ini_len}`).getDisplayValues();
  

  let cur_date = new Date();
  const lastWeekDate = new Date(cur_date.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  data = data.filter((item) => {
    let row_date = new Date(ddmmyyyy_to_yyyymmdd(item[date_idx]));
    return lastWeekDate < row_date;
  });

  // clearing 

  sheet.getRange(`A2:${last_col_char}${ini_len}`).clear();
  sheet.getRange(`A2:${last_col_char}${data.length+1}`).setValues(data); 
}

function clear_reportsHelpingMetaData(){
  clean_up_trigger("ReportsHelpingMetaData" , 0);
}

// -4200086864 : btesting
function generate_doubts(type = "single" , row = null ){
  let ui = SpreadsheetApp.getUi();
  if (type === "single" && SpreadsheetApp.getActiveSheet().getName() !== "Doubts") {
    ui.alert("Please use this option inside Doubts sheet");
    return;
  }

  let e = "uc";
  if(e === "uc"){
    ui.alert("🚧Under Construction!🚧 - This feature will be available soon.");
    return;
  }

  let code_to_subject = {"C" : "Chemistry" , "P" : "Physics" , "M" : "Maths"};
  let zoom_accounts_data = ss.getSheetByName("Zoom Accounts").getDataRange().getDisplayValues();
  let telegram_data = ss.getSheetByName("TelegramGroups").getDataRange().getDisplayValues();
  let msg_template_data = ss.getSheetByName("Message").getDataRange().getValues();
  let doubts = ss.getSheetByName("Doubts");

  for (let i = 0; i < telegram_data.length; i++) {
    telegram_maps_chatid[telegram_data[i][0]] = telegram_data[i][1];
  }

  let zoom_f_c_maps_credentials = {};
  for (let i = 1; i < zoom_accounts_data.length; i++) {
    zoom_f_c_maps_credentials[zoom_accounts_data[i][0]] = [
      zoom_accounts_data[i][4],
      zoom_accounts_data[i][5],
      zoom_accounts_data[i][6],
      zoom_accounts_data[i][1]
    ];
  }

  
  let cell = doubts.getActiveCell();
  if (row === null && type === "single") {
    row = cell.getRowIndex();
  }
  let no_of_slots = doubts.getRange(`B${row}`).getValue();
  if (type === "single" && (cell.getValue().toString().trim().length == 0 || row == 1 || no_of_slots.toString().trim().length == 0)) {
    ui.alert("Please select a starting row of a doubt collection");
    return ;
  }

  let final_msg = msg_template_data[2][1] + "\n\n";
  let telegram_grps_chat_ids = [];
  let doubts_data = doubts.getRange(`A${row}:F${row+no_of_slots-1}`).getDisplayValues();

  for(let r=0 ; r<no_of_slots ; r++){
    let slot_msg = "";
    let faculty = doubts_data[r][4].toString().split("|~|");
    let tel_groups = doubts_data[r][5].toString().split("|~|");

    if(tel_groups.length){
      for(let i=0 ; i<tel_groups.length ; i++){
        if(telegram_maps_chatid[tel_groups[i]] && !isNaN(telegram_maps_chatid[tel_groups[i]]) && !telegram_grps_chat_ids.includes(telegram_maps_chatid[tel_groups[i]])){
          telegram_grps_chat_ids.push(telegram_maps_chatid[tel_groups[i]]);
        }
      }
    }


    if(faculty.length){
      slot_msg = `<blockquote>\nZoom (ONLINE)-SLOT ${r+1}  (${tConvert(doubts_data[r][2])} to ${tConvert(doubts_data[r][3])})\n\n`;
      for(let i=0 ; i<faculty.length ; i++){
        if(zoom_f_c_maps_credentials[faculty[i]]){
          let client_id = zoom_f_c_maps_credentials[faculty[i]][0] ;
          let client_secret = zoom_f_c_maps_credentials[faculty[i]][1];
          let account_id = zoom_f_c_maps_credentials[faculty[i]][2];

          let access_token_res = get_zoom_access_token(account_id , client_id , client_secret);

          let created_meeting_data;

          if (access_token_res) {
            try {
              let body = {
                "agenda": `Doubts Session by ${faculty[i]} sir from ${tConvert(doubts_data[r][2])} to ${tConvert(doubts_data[r][3])} || ${doubts_data[0][0]}`,
                "default_password": false,
                "duration": cal_duration(doubts_data[r][2], doubts_data[r][3]),
                "password": "987",
                "pre_schedule": false,
                "schedule_for": zoom_f_c_maps_credentials[faculty[i]][3],
                "settings": {
                  "allow_multiple_devices": false,
                  "approval_type": 2,
                  "audio": "voip",
                  "authentication_domains": "bakliwaltutorials.com,bakliwaltutorialsiit.onmicrosoft.com",
                  "auto_recording": "cloud",
                  "contact_email": zoom_f_c_maps_credentials[faculty[i]][3],
                  "contact_name": "BT-Lecture",
                  "email_notification": true,
                  "encryption_type": "enhanced_encryption",
                  "focus_mode": true,
                  "host_video": false,
                  "jbh_time": 0,
                  "join_before_host": false,
                  "meeting_authentication": true,
                  "mute_upon_entry": true,
                  "participant_video": false,
                  "use_pmi": false,
                  "waiting_room": false,
                  "watermark": false,
                  "internal_meeting": false,
                  "continuous_meeting_chat": {
                    "enable": true,
                    "auto_add_invited_external_users": false
                  },
                  "participant_focused_meeting": false,
                  "auto_start_meeting_summary": false,
                  "auto_start_ai_companion_questions": false
                },
                "start_time": `${ddmmyyyy_to_yyyymmdd(doubts_data[0][1])}T${doubts_data[r][2]}`,
                "timezone": "Asia/Calcutta",
                "topic": `Doubts Session by ${faculty[i]} sir from ${tConvert(doubts_data[r][2])} to ${tConvert(doubts_data[r][3])} || ${doubts_data[0][0]}`,
                "type": 2
              }

              let created_meeting_data_res = UrlFetchApp.fetch(`https://api.zoom.us/v2/users/${zoom_f_c_maps_credentials[faculty[i]][3]}/meetings`, {
                "method": "POST",
                "headers": {
                  'Content-Type': 'application/json',
                  'Authorization': `${access_token_res.token_type} ${access_token_res.access_token}`
                },
                "payload": JSON.stringify(body),
              });

              if (created_meeting_data_res.getResponseCode() === 429) {
                slot_msg += `${faculty[i]} : Too Many Requests\n`;
                continue;
              }
              if (created_meeting_data_res.getResponseCode() === 400) {
                slot_msg += `${faculty[i]} : Bad Request\n`;
                continue;
              }
              if (created_meeting_data_res.getResponseCode() === 404) {
                slot_msg += `${faculty[i]} : Not Found, (User does not exist: {userId}.)\n`;
                continue;
              }

              created_meeting_data = JSON.parse(created_meeting_data_res);
              let meet_credentials = created_meeting_data.join_url + '\n\nMeeting ID: ' + created_meeting_data.id + '\nPasscode: ' + created_meeting_data.password;

              slot_msg += `${code_to_subject[faculty[i].substring(0 , 1)]} : ${faculty[i]} Sir ${tConvert(doubts_data[r][2])} to ${tConvert(doubts_data[r][3])} \n ${meet_credentials} \n`;

            } catch (e) {
              slot_msg += `${faculty[i]} : ${e.message} \n`;
            }
          }else{
            slot_msg += `${faculty[i]} : error in generating Access token \n`;
          }

        }else{
          slot_msg += `${faculty[i]} : zoom credentials does not exits for this faculty code. \n`;
        }
      }

      slot_msg += `</blockquote>\n\n`;
    }

    final_msg += slot_msg;
  }

  final_msg += "\n\n"+ msg_template_data[2][2];
  final_msg.replace("[DATE]" , doubts_data[0][0]);
  doubts.getRange(`G${row}`).setValue(final_msg);

  let link_preview_options = JSON.stringify({"is_disabled":true});
  if(telegram_grps_chat_ids.length){
    let token = ss.getSheetByName("Information").getRange("A6").getValue();
    for(let i=0 ; i<telegram_grps_chat_ids.length ; i++){
      let chat_id = telegram_grps_chat_ids[i];
      final_msg = encodeURI(final_msg);

      try {
        let url = `https://api.telegram.org/bot${token}/sendMessage?chat_id=${chat_id}&text=${final_msg}&parse_mode=HTML&link_preview_options=${link_preview_options}`;
        let res_msg = JSON.parse(UrlFetchApp.fetch(url));
        if (res_msg.ok == true) {
          doubts.getRange(`H${row}`).setValue("messages sent!!");
        } else {
          if (res_msg.description) {
            doubts.getRange(`H${row}`).setValue(res_msg.description);
          }
          else {
            doubts.getRange(`H${row}`).setValue("Some error occured");
          }
        }

      } catch (e) {
        doubts.getRange(`H${row}`).setValue(e.message);
      }
    }
  }else{
    doubts.getRange(`H${row}`).setValue("No telegram group is selected");
  }

}

function bulk_generate_doubts(e = "uc"){
  let ui = SpreadsheetApp.getUi();

  if (e === "uc") {
    ui.alert("🚧Under Construction!🚧 - This feature will be available soon.");
    return;
  }

  if (SpreadsheetApp.getActiveSheet().getName() !== "Doubts") {
    ui.alert("Please use this option inside Doubts sheet");
    return;
  }

  let input = ui.prompt("Please enter the range (example: 3-30)");

  let button = input.getSelectedButton();
  let doubt_sheet = ss.getSheetByName("Doubts");

  if (button === ui.Button.OK) {
    let range = input.getResponseText();
    let range_arr = range.split("-");
    if (range_arr.length !== 2) {
      ui.alert("Enter the valid range");
      return;
    }

    if (isNaN(range_arr[0].toString().trim()) || isNaN(range_arr[1].toString().trim())) {
      ui.alert("Enter the valid range");
      return;
    }

    let start = Number(range_arr[0]);
    let end = Number(range_arr[1]);

    if (start === 1 || end === 1) {
      ui.alert("please do not enter value 1 (i.e. Header)");
      return;
    }
    if( doubt_sheet.getRange(`B${start}`).getValue().toString().trim().length() === 0 ||
        doubt_sheet.getRange(`B${end}`).getValue().toString().trim().length() === 0 ){
          ui.alert("Please Enter the starting row of an doubts collection.")
          return ;
    }

    if (end < start) {
      ui.alert("Start should be lesser than end");
      return;
    }

    if (end - start + 1 > 50) {
      ui.alert("The range should be smaller than 50.");
      return;
    }

    let all_done = true;
    let collection_len = null;
    for (let row = start; row <= end; row+=collection_len) {
      generate_doubts("bulk", row);
      collection_len = Number(doubt_sheet.getRange(`B${row}`).getValue());
    }

    if (!all_done) {
      ui.alert("Some messages are not sent kindly check.");
      return;
    }


  } else if (button === ui.Button.CLOSE) {
    Logger.log("The user clicked the [X] button and closed the prompt dialog.");
  }
}
