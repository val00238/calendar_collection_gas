function doGet(e) {
  t = HtmlService.createTemplateFromFile('index.html');
  t.title = 'CalendarCollection';
  t.sdate = getStartDateString();
  t.edate = getEndDateString();

  return t.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getStartDateString() {
  var date = new Date();
  date .setDate(1);

  var date_string = [
    date.getFullYear(),
    ("0"+(date.getMonth()+1)).slice(-2),
    ("0"+date.getDate()).slice(-2)
  ].join( '-' );

 return date_string;
}

function getEndDateString() {
  var date = new Date();
  date.setDate(1);
  date.setMonth(date.getMonth() + 1);
  date.setDate(0);

  var date_string = [
    date.getFullYear(),
    ("0"+(date.getMonth()+1)).slice(-2),
    ("0"+date.getDate()).slice(-2)
  ].join( '-' );

 return date_string;
}

function toLocaleDateString(date) {
  var date_string = [
    date.getFullYear(),
    ("0"+(date.getMonth()+1)).slice(-2),
    ("0"+date.getDate()).slice(-2)
  ].join( '/' );

  var time_string = [
    ("0"+(date.getHours())).slice(-2),
    ("0"+(date.getMinutes())).slice(-2),
    ("0"+(date.getSeconds())).slice(-2)
  ].join( ':' );

  return date_string + " " + time_string;
}

function processForm(form) {
  var s_date = form.sdate;
  var e_date = form.edate;
  var bd_state = form.bd;

  var s_date_f = s_date.replace(/-/g,"/");
  var e_date_f = e_date.replace(/-/g,"/");

  var elapsed_time = new Date(e_date_f) - new Date(s_date_f);
  if (elapsed_time <= 0) {
    return "<div class=\"alert alert-danger\">The designation in a period is wrong, so please check it ! <strong>'" + s_date + " - " + e_date + "'</strong></div>";
  }

  var cal_events = getCalendarEvents(s_date_f, e_date_f);
  if (cal_events.length <= 0) {
    return "<div class=\"alert alert-danger\">A schedule wasn't registered with your calendar a period of <strong>'" + s_date + " - " + e_date + "'</strong></div>";
  }

  var url = createSpreadsheetFile(s_date + "_" + e_date);
  var spread_sheet = SpreadsheetApp.openByUrl(url);

  var event_array = getEventData(cal_events);

  var count = writeSpreadsheet(event_array, spread_sheet);
  if (bd_state) {
    var count_bd = writeBDFormatSheet(event_array, spread_sheet);
  }

  result = createResponse(event_array, url);

  return result.join("");
}

function createResponse(event_array, url) {
  var response_array = [];

  if (url) {
    response_array.push('<form class="form-horizontal" onsubmit="javascript:return false;" id="response_form">');
    response_array.push('<fieldset>');
    response_array.push('<div class="form-group">');
    response_array.push('<div class="col col-xs-1 col-sm-1">');
    response_array.push('<button type="open" class="btn  btn-primary btn-lg" onclick="window.open(\'' + url + '\');"> open sheet </button>');
    response_array.push('</div>');
    response_array.push('</div>');
    response_array.push('</fieldset>');
    response_array.push('</form>');
  }

  response_array.push('<table class="table table-striped table-bordered table-hover">');
  response_array.push('<thead>');

  // add header
  var header = createHeader(event_array[0]);
  for (var i=0; i < header.length; i++) {
    response_array.push('<th class="head">');
    response_array.push(header[i]);
    response_array.push('</th>');
  }

  response_array.push('</thead>');
  response_array.push('<tbody>');

  // add data
  var array_keys = Object.keys(event_array[0]);
  for (var i=0; i < event_array.length; i++) {
    response_array.push('<tr>');

    var event_obj = event_array[i];
    var event_keys = Object.keys(event_obj);

    var write_array = [];
    for (var n=0; n < array_keys.length; n++) {
      var item = event_obj[event_keys[n]];

      response_array.push('<td>');

      if (isType("Array", item)) {
        for (var k=0; k < item.length; k++) {
          if (item[k].indexOf("http") >= 0) {
            response_array.push('<a href="' + item[k] + '" target="_new">' + item[k] + '</a><br>');
          } else {
            response_array.push(item[k]);
          }
        }
      } else {
        response_array.push(item);
      }

      response_array.push('</td>');
    }
    response_array.push('</tr>');
  }

  response_array.push('</tbody>');
  response_array.push('</table>');

  return response_array;
}

function getEventData(calendar_events) {
  var event_array = [];

  for(var i=0; i < calendar_events.length; i++) {
    var event = calendar_events[i];

    var event_obj = getEventObject(event);
    var event_days = getEventSpan(event);
    if (event_obj) {
      event_array.push(event_obj);
    }

    // event span
    for (var n=1; n < event_days; n++) {
      var span_date = event.getStartTime();
      var day_of_month = span_date.getDate();
      span_date.setDate(day_of_month + n);

      var event_span_obj = getEventObject(event, span_date);
      event_array.push(event_span_obj);
    }
  }

  // sort
  event_array.sort (
    function(a,b){
      var a_val = a["date"];
      var b_val = b["date"];
      if( a_val < b_val ) return -1;
      if( a_val > b_val ) return 1;
      return 0;
    }
  );

  return event_array;
}

function getEventSpan(event) {
  // Set the unit values in milliseconds.
  var msec_per_minute = 1000 * 60;
  var msec_per_hour = msec_per_minute * 60;
  var msec_per_day = msec_per_hour * 24;

  var date_msec = event.getEndTime().getTime();

  // Get the difference in milliseconds.
  var interval = date_msec - event.getStartTime().getTime();

  // Calculate how many days the interval contains. Subtract that
  // many days from the interval to determine the remainder.
  var days = Math.floor(interval / msec_per_day );

  return days;
}

function getEventObject(event, event_date) {
  var day_event = {};

  day_event.date = event.getStartTime();
  if (event_date) {
    day_event.date = event_date;
  }
  day_event.title = event.getTitle();
  day_event.location = event.getLocation();
  day_event.description = event.getDescription();
  day_event.startTime = event.getStartTime();
  day_event.endTime = event.getEndTime();
  day_event.url = "";
  day_event.type = "";
  day_event.created = event.getDateCreated();

  var des_text = day_event.description;
  var url_regexp = new RegExp("https?://[a-zA-Z0-9\-_\.:@!~*'\(¥);/?&=\+$,%#]+", "g");
  var url_text = des_text.match(url_regexp);

  if (url_text) {
    day_event.url = url_text;
  }

  var type_regexp = new RegExp("(＜|<).+(＞|>)", "g");
  var type_text = des_text.match(type_regexp);

  if (type_text) {
    day_event.type = type_text;
  }

  return day_event;
}

function createHeader(event_obj) {
  var obj_keys = Object.keys(event_obj);

  var header_array = [];
  for (var i=0; i < obj_keys.length; i++) {
    var key_name = obj_keys[i];

    var header_text = "";
    switch (key_name) {
      case "date":
        header_text = "日付";
        break;
      case "title":
        header_text = "タイトル";
        break;
      case "location":
        header_text = "場所";
        break;
      case "description":
        header_text = "説明";
        break;
      case "url":
        header_text = "関連URL";
        break;
      case "type":
        header_text = "種別";
        break;
      case "startTime":
        header_text = "開始日";
        break;
      case "endTime":
        header_text = "終了日";
        break;
      case "created":
        header_text = "作成日";
        break;
    }

    header_array.push(header_text);
  }

  return header_array;
}

function writeSpreadsheet(event_array, spread_sheet) {
  var array_keys = Object.keys(event_array[0]);

  var max_column = event_array.length;
  var max_row    = array_keys.length;
  var start_column = 1;
  var start_row = 1;

  var sheets = spread_sheet.getSheets();
  if (sheets.length <= 0) {
    return;
  }
  var sheet_obj = sheets[0];
  sheet_obj = sheet_obj.clear();

  for (var i=0; i < event_array.length; i++) {
    var event_obj = event_array[i];
    var event_keys = Object.keys(event_obj);

    var write_array = [];
    for (var n=0; n < array_keys.length; n++) {
      var item = event_obj[event_keys[n]];

      if (isType("Array", item)) {
        item = item.join(String.fromCharCode(10));
      }
      //if (isType("Date", item)) {
      //  item = toLocaleDateString(item);
      //}

      write_array.push(item);
    }

    var range = sheet_obj.getRange(start_row+i, start_column, 1, write_array.length);
    if (i == 0) {
      var header = createHeader(event_obj);
      range.setValues( [header] );
      range.setBackground("#CCCCCC");
      start_row++;
      range = sheet_obj.getRange(start_row+i, start_column, 1, write_array.length);
    }
    range.setValues( [write_array] );
  }

  return max_column;
}

function writeBDFormatSheet(event_array, spread_sheet) {
  var header = ["日付", "タイトル", "内容", "参加者", "URL", "備考", "種別"];
  var array_keys = Object.keys(event_array[0]);
  var max_column = event_array.length;
  var max_row    = header.length;
  var start_column = 1;
  var start_row = 1;

  var sheet_obj = spread_sheet.insertSheet("BD_Format");
  sheet_obj = sheet_obj.clear();

  for (var i=0; i < event_array.length; i++) {
    var event_obj = event_array[i];
    var event_keys = Object.keys(event_obj);

    var item_date = "";
    var item_title = "";
    var item_description = "";
    var item_entry = "";
    var item_url = "";
    var item_recital = "";
    var item_type = "";
    for (var n=0; n < array_keys.length; n++) {
      var key_name = event_keys[n];
      var item = event_obj[key_name];

      if (isType("Array", item)) {
        item = item.join(String.fromCharCode(10));
      }

      switch (key_name) {
        case "date":
          item_date = item;
          break;
        case "title":
          item_title = item;
          break;
        case "description":
          var url_regexp = new RegExp("https?://[a-zA-Z0-9\-_\.:@!~*'\(¥);/?&=\+$,%#]+", "g");
          var type_regexp = new RegExp("(＜|<).+(＞|>)", "g");

          item_description = item.replace(url_regexp, "");
          item_description = item_description.replace(type_regexp, "");
          break;
        case "url":
          item_url = item;
          break;
        case "type":
          item_type = item;
          break;
      }
    }

    // same line as header
    var write_array = [item_date, item_title, item_description, item_entry, item_url, item_recital, item_type];

    var range = sheet_obj.getRange(start_row+i, start_column, 1, write_array.length);
    if (i == 0) {
      range.setValues( [header] );
      range.setBackground("#CCCCCC");
      start_row++;
      range = sheet_obj.getRange(start_row+i, start_column, 1, write_array.length);
    }

    range.setValues( [write_array] );
  }

  return max_column;
}

function isType(type, obj) {
  var clas = Object.prototype.toString.call(obj).slice(8, -1);

  return obj !== undefined && obj !== null && clas === type;
}

function getCalendarEvents(s_date, e_date) {
  var user_mail = Session.getActiveUser().getEmail();

  var cal = CalendarApp.getCalendarById(user_mail);

  var s_date_time = new Date(s_date + " 00:00:00");
  var e_date_time = new Date(e_date + " 00:00:00");
  e_date_time.setDate( e_date_time.getDate() + 1 );

  var events = cal.getEvents(s_date_time, e_date_time);

  return events
}

var SHEET_FOLDER_NAME = "_sheets";
function createSpreadsheetFile(time_span) {
  var folder_name = SHEET_FOLDER_NAME;
  var folders = DriveApp.getFoldersByName(folder_name);

  var folder;
  while (folders.hasNext()) {
    folder = folders.next();
  }

  if (folder == null) {
    folder = DriveApp.createFolder(folder_name);
  }

  var dateString = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss_');
  var file_anme = dateString + time_span;

  var ss_id = SpreadsheetApp.create(file_anme).getId();
  var ss_file = DriveApp.getFileById(ss_id);
  var url = ss_file.getUrl();

  folder.addFile(ss_file);
  DriveApp.getRootFolder().removeFile(ss_file);

  return url;
}
