var num_first_fields_required = 7; // I primi 7 campi sono quelli obbligatori: ['Report Name'], ['Account ID'], ['Property ID'], ['View ID'], ['Start Date'], ['End Date'], ['Metrics']
var report_configuration_name = "Report Configuration 360";
var html_file_name_ref = "html_file_name";
var arr_fields = [['Report Name*'], ['Account ID*'], ['Property ID*'], ['View ID*'], ['Start Date*'], ['End Date*'], ['Metrics*'], ['Dimensions'], ['Filters'], ['Segments'], ['Creation date'], ['Last update'], ['Status'], ['Report ID'], ['Document ID'], ['Next scheduled check']];
var document_trigger_properties_name = 'gaur_default_trigger_hourly';
var document_schedule_properties_name = 'gaur_scheduled_trigger_details'; // fake trigger: sfrutta il trigger orario e verifica le impostazioni di schedulazione per rinizializzare i report
var document_autorerun_properties_name = 'gaur_auto_rerun_to_complete';

var status_pending = 'PENDING';
var status_completed = 'COMPLETED';
var status_invalid = 'INVALID REQUEST';
var status_incomplete = 'Incomplete configuration!';
var status_error = 'Configuration error!';
var status_duplicate = 'Duplicate name!';

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
  .addItem('Create new report', 'newReport')
  .addItem('Run reports', 'runReportsLoader')
  .addItem('Schedule reports', 'scheduleReports')
  .addToUi();
}

function onEdit(e) {
  var activeSheet = e.source.getActiveSheet();
  var activeSheetName = activeSheet.getName();
  var range_selected = e.range.getA1Notation(); // potrebbe essere D3 ma anche D3:D8 oppure D3:F6
  
  var arr_range_selected = range_selected.split(":");
  var arr_range_selected_len = arr_range_selected.length;
    
  var cell_column_letter_first = arr_range_selected[0].replace(/[0-9]/g, '');
  var cell_column_letter_second = cell_column_letter_first;
  if (arr_range_selected_len > 1) { cell_column_letter_second = arr_range_selected[1].replace(/[0-9]/g, ''); }
  
  var cell_column_number_first = arr_range_selected[0].replace(/[A-Z]/g, '');
  var cell_column_number_second = cell_column_number_first;
  if (arr_range_selected_len > 1) { cell_column_number_second = arr_range_selected[1].replace(/[A-Z]/g, ''); }

  if (activeSheetName == report_configuration_name) {
    if (cell_column_letter_first != "A" && parseInt(cell_column_number_second) < 12) {
      activeSheet.getRange(cell_column_letter_first+"12:"+cell_column_letter_second+"17").clearContent(); // Se, ad esempio, si modificano i dettagli del report nella cella B viene cancellato automaticamente il contenuto da B12 a B16 (perché ipoteticamente si tratterà di una chiamata ad un nuovo report) così come quelli di un range (es: B e C insieme)
    }
  }
}

function scheduleReports() {
  modal_title = "Schedule reports";
  modal_html_name = "modal_report_schedule";
  modal_width = 700;
  modal_height = 160;
  showModalDialogCustom(modal_width,modal_height,modal_html_name,modal_title);
}

function provideHTMLContent() {
  var cache = CacheService.getDocumentCache();
  var html_file_name_cached = cache.get(html_file_name_ref);
  return HtmlService.createHtmlOutputFromFile(html_file_name_cached).getContent();
}

function newReport() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar_report_new').setTitle('Create a new unsampled report');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function runReports() {
  var documentProperties = PropertiesService.getDocumentProperties();

  documentProperties.setProperty(document_autorerun_properties_name, '1'); // valorizzo la property per indicare preventivamente la necessità di effettuare il run automatico per i report pending (successivamente se non ci sono report pending la svuoto, in questo modo il rerun automatico non entra più in questa funzione)

  var document_trigger_properties_name_value = documentProperties.getProperty(document_trigger_properties_name);
  if (!document_trigger_properties_name_value) {
    delete_ALLTriggers_ALLProperties(); // elimina tutti i trigger e tutte le properties in modo da ripartire da situazione pulita anche per cuhi ha installato il vecchio add-on. In teoria di qui ci passa solo la prima volta che si esegue il primo report.
    setDefaultTrigger(documentProperties); // imposto la prima volta il trigger orario di default
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss_sheet = ss.getSheetByName(report_configuration_name);

  //const TIME_TO_WAIT_FOR_TRIGGER_ATTIVATION = 3630000; // 60 minuti (e 30 secondi, per garantire l'intervallo di 1h ed evitare l'errore "gli eventi orologio devono essere pianificati con intervalli di almeno 1 ora l'uno dall'altro")
  var report_name = '';
  var select_account = '';
  var select_property = '';
  var select_view = '';
  var date_begin = '';
  var date_end = '';  
  var report_m = '';
  var report_d = '';
  var report_f = ''; 
  var report_s = '';
  
  var report_creation = '';
  var report_update = '';
  var report_status = '';
  var report_id = '';
  var report_document_id = '';
  
  var lastColumn = ss_sheet.getLastColumn();
  var arr_fields_len = arr_fields.length;
  
  var created_row = arr_fields_len - 4;
  var updated_row = arr_fields_len - 3;
  var status_row = arr_fields_len - 2;
  var report_id_row = arr_fields_len - 1;
  var report_docid_row = arr_fields_len;
  var report_next_chk = arr_fields_len + 1;
  
  var bln_duplicate_name = false;
  var bln_invalid_reports = false;
  var arr_sheetname = [];
  var arr_sheetname_len = 0;
  
  var report_run = 0;
  var report_run_not = 0;

  var counter_done = 0;
  var counter_running = 0;
  var counter_error = 0;
  var counter_duplicate = 0;
  
  var tot_ok = 0;
  var tot_ko = 0;

  var string_for_check_appoggio = '';
  var arr_csv_data_insert = [];
  var cnt_csv = 0;
  var return_data = false; // utilizzato solo per recuperare il return true da importReportFromDrive (boolean poi non utilizzato)
  
  //var dateObj_for_trigger = Date.now() + TIME_TO_WAIT_FOR_TRIGGER_ATTIVATION; // next time to check for trigger and display in field 16
  //var report_next_check = Utilities.formatDate(new Date(dateObj_for_trigger), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd hh:mm a"); // format date for display next time to check in field 16
  var report_next_check = 'every hour until completion';

  for (var i=2; i<=lastColumn; i++) {
    bln_duplicate_name = false;
    var resource = {};
    var counter_empty = 0;
    var current_report_options = ss_sheet.getRange(2,i,arr_fields_len,1).getValues();
    var arr_values_sheet_details = [];
    var error_details = 'No error';
    var current_status = '';
    var bln_report_done = false;
    
    for (var j=0; j<arr_fields_len;j++) {
      if (current_report_options[j] == '' || current_report_options[j] == ' ') {
        if (j < num_first_fields_required) { counter_empty = counter_empty + 1; } // i primi 7 campi sono obbligatori
      }
      if (j == 0) { report_name = current_report_options[j] }; // required
      if (j == 1) { select_account = current_report_options[j] }; // required
      if (j == 2) { select_property = current_report_options[j] }; // required
      if (j == 3) { select_view = current_report_options[j] }; // required
      if (j == 4) { date_begin = current_report_options[j] }; // required
      if (j == 5) { date_end = current_report_options[j] }; // required
      if (j == 6) { report_m = current_report_options[j] }; // required
      if (j == 7) { report_d = current_report_options[j] };
      if (j == 8) { report_f = current_report_options[j] };
      if (j == 9) { report_s = current_report_options[j] };
      if (j == 10) { report_creation = current_report_options[j] }; // Creation date
      if (j == 11) { report_update = current_report_options[j] }; // Last update
      if (j == 12) { report_status = current_report_options[j] }; // Status
      if (j == 13) { report_id = current_report_options[j] }; // Report ID
      if (j == 14) { report_document_id = current_report_options[j] }; // Document ID
      //if (j == 15) { report_next_check = current_report_options[j] }; // Next scheduled check
    }
    
    if (counter_empty != num_first_fields_required) {  // SE E' UGUALE SIGNIFICA CHE TUTTI I CAMPI OBBLIGATORI DELLA COLONNA (I PRIMI 7 NEL CASO SPECIFICO, vedere variabile num_first_fields_required) SONO VUOTI PER CUI SALTA ALLA COLONNA SUCCESSIVA SENZA CONSIDERARE QUELLA CORRENTE
      if (counter_empty == 0) {
        arr_sheetname_len = arr_sheetname.length;
        for (var y=0; y<arr_sheetname_len; y++) {
          if (arr_sheetname[y]+'' == report_name+'') {
            bln_duplicate_name = true;
            break;
          }
        }

        if (!bln_duplicate_name) {
          if (report_status == "" || report_status == status_invalid) {
            date_begin = validateReportDate(date_begin);
            date_end = validateReportDate(date_end);
            if (report_name == "" || select_account == "" || select_property == "" || select_view == "" || date_begin == "" || date_end == "" || report_m == "") {
              bln_invalid_reports = true;
              ss_sheet.getRange(status_row,i,1,1).setValue("INVALID REQUEST"); // LA MODAL DI RIEPILOGO DOVRA' INFORMARE L'UTENTE CHE CI SONO REPORT NON VALIDI E QUALI CAMPI VERIFICARE PERCHE' SONO OBBLIGATORI
              counter_error = counter_error + 1;
              arr_sheetname.push(report_name);
              SpreadsheetApp.flush();
              Utilities.sleep(200);
              continue;
            }
          }

          if (report_status == status_pending || report_status == status_completed) {
            string_for_check_appoggio = get_UnsampledReport(select_account, select_property, select_view, report_id);
            // only for test
            //string_for_check_appoggio = {"accountId":"160739218", "kind":"analytics#unsampledReport", "created":"2020-12-14T13:32:22.382Z", "profileId":"213646178", "downloadType":"GOOGLE_DRIVE", "id":"lbvlO__rT2GgvCGS4pxlxg", "title":"Prova test 123", "updated":"2020-12-14T13:32:29.874Z", "driveDownloadDetails":{"documentId":"1fWtvHunH6yGYpyHUanbfg-TUH-im8vPz"}, "webPropertyId":"UA-160739218-1", "status":"COMPLETED", "selfLink":"https://www.googleapis.com/analytics/v3/management/accounts/160739218/webproperties/UA-160739218-1/profiles/213646178/unsampledReports/lbvlO__rT2GgvCGS4pxlxg"}
            if (string_for_check_appoggio.status == status_completed) {
              Logger.log(string_for_check_appoggio);
              ss_sheet.getRange(created_row,i,6,1).setValues([[string_for_check_appoggio.created], [string_for_check_appoggio.updated], [string_for_check_appoggio.status], [string_for_check_appoggio.id], [string_for_check_appoggio.driveDownloadDetails.documentId], ['-']]); // popola questi 6 campi: ['Creation date'], ['Last update'], ['Status'], ['Report ID'], ['Document ID'], ['Next scheduled check']
              return_data = importReportFromDrive(ss, string_for_check_appoggio.title, string_for_check_appoggio.driveDownloadDetails.documentId);
              counter_done = counter_done + 1;
              report_run = report_run + 1;
              arr_sheetname.push(report_name);
              SpreadsheetApp.flush();
              Utilities.sleep(200);
              continue;
            }
            ss_sheet.getRange(report_next_chk,i,1,1).setValues([[report_next_check]]);
            counter_running = counter_running + 1;
            report_run = report_run + 1;
            arr_sheetname.push(report_name);
            SpreadsheetApp.flush();
            Utilities.sleep(200);
            continue;
          }

          // Fix alt+enter to separate metrics and dimensions
          report_m = String(report_m).replace(/ /g, ",");
          report_m = report_m.replace(/\n/g, ",");
          report_m = report_m.replace(/,,/g, ",");
          
          report_d = String(report_d).replace(/ /g, ",");
          report_d = report_d.replace(/\n/g, ",");
          report_d = report_d.replace(/,,/g, ",");
          
          resource = {
            'title': report_name+"",
            'start-date': date_begin+"",
            'end-date': date_end+"",
            'metrics': report_m+""
          };
          
          if (report_d != '') { resource["dimensions"] = report_d+""; }
          if (report_f != '') { resource["filters"] = report_f+""; }
          if (report_s != '') { resource["segment"] = report_s+""; }
          
          string_for_check_appoggio = insert_UnsampledReport(resource, select_account, select_property, select_view);
          // only for test
          //string_for_check_appoggio = {"accountId":"160739218", "kind":"analytics#unsampledReport", "created":"2020-12-14T13:33:55.844Z", "profileId":"213646178", "id":"8ATLgDzFSOqMpPWKulQb_w", "title":"Prova test 123", "updated":"2020-12-14T13:33:57.733Z", "webPropertyId":"UA-160739218-1", "status":"PENDING", "selfLink":"https://www.googleapis.com/analytics/v3/management/accounts/160739218/webproperties/UA-160739218-1/profiles/213646178/unsampledReports/8ATLgDzFSOqMpPWKulQb_w"}
          
          //browserMsgBox(string_for_check_appoggio)
          
          if (string_for_check_appoggio) {
            ss_sheet.getRange(created_row,i,6,1).setValues([[string_for_check_appoggio.created], [string_for_check_appoggio.updated], [string_for_check_appoggio.status], [string_for_check_appoggio.id], [''], [report_next_check]]);
          }
          
          counter_running = counter_running + 1;
          report_run = report_run + 1;
          
        } else {
          counter_duplicate = counter_duplicate + 1;
          ss_sheet.getRange(status_row,i,1,1).setValue(status_duplicate);
        }
      } else {
        report_run_not = report_run_not + 1;
        ss_sheet.getRange(status_row,i,1,1).setValue(status_incomplete);
      }
      arr_sheetname.push(report_name);
    }
    SpreadsheetApp.flush();
    Utilities.sleep(200);
  }
  
  var tot_ok = counter_done + counter_running;
  var tot_ko = report_run_not + counter_error + counter_duplicate;
  var htmlOutput = "<p style='margin: 0;'><span [%CLASS_GREEN%]>Reports run <strong>[%REPORT_RUN%]</strong></span>:<br />[%REPORT_DONE%] reports '" + status_completed + "'<br />[%REPORT_RUNNING%] reports '" + status_pending + "'[%REPORT_RERUN%]</p>";
  htmlOutput = htmlOutput + "<p style='margin: 14px 0 22px;'><span [%CLASS_RED%]>Reports <u>not</u> run <strong>[%REPORT_NOT_RUN%]</strong></span>:<br />[%REPORT_ERROR%] reports '" + status_error + "'<br />[%REPORT_INCOMPLETE%] reports '" + status_incomplete + "'<br />[%REPORT_DUPLICATE%] reports '" + status_duplicate + "'</p>";
  htmlOutput = htmlOutput + "<button onclick='google.script.host.close();'>OK</button>"
  htmlOutput = htmlOutput.replace("[%REPORT_RUN%]",report_run);
  htmlOutput = htmlOutput.replace("[%REPORT_DONE%]",counter_done);
  htmlOutput = htmlOutput.replace("[%REPORT_RUNNING%]",counter_running);
  htmlOutput = htmlOutput.replace("[%REPORT_NOT_RUN%]",report_run_not + counter_error + counter_duplicate);
  htmlOutput = htmlOutput.replace("[%REPORT_ERROR%]",counter_error);
  htmlOutput = htmlOutput.replace("[%REPORT_INCOMPLETE%]",report_run_not);
  htmlOutput = htmlOutput.replace("[%REPORT_DUPLICATE%]",counter_duplicate); 
  if (counter_running > 0) {
    //createTriggerBasedTime(documentProperties, script_trigger_name_base, 'runReports', dateObj_for_trigger);
    //htmlOutput = htmlOutput.replace("[%REPORT_RERUN%]","<br /><span class='help'>The script will re-run the reports automatically and silently every <strong>" + millisToMinutesAndSeconds(TIME_TO_WAIT_FOR_TRIGGER_ATTIVATION, false)  + " minutes</strong> to check the status of the pending reports until the result is obtained and shown.<br />You can force execution by choosing <strong>Add-ons » GA360 Unsampled » Run reports</strong>.<span>");
    htmlOutput = htmlOutput.replace("[%REPORT_RERUN%]","<br /><span class='help'>The script will re-run the reports automatically and silently hourly to check the status of the pending reports until the result is obtained and shown.<br />You can force execution by choosing <strong>Add-ons » GA360 Unsampled » Run reports</strong>.<span>");
  } else { // se invece non ci sono report pending
    documentProperties.deleteProperty(document_autorerun_properties_name); // se non ci sono report pending elimino la property che mi indica la necessità di rerun della funzione in questione: runReports()
  }
  htmlOutput = htmlOutput.replace("[%REPORT_RERUN%]","");
  
  if (tot_ok > 0) { htmlOutput = htmlOutput.replace("[%CLASS_GREEN%]","style='color:#1e9c5a;'"); }
  if (tot_ko > 0) { htmlOutput = htmlOutput.replace("[%CLASS_RED%]","style='color:#db4c3f;'"); }
  htmlOutput = htmlOutput.replace("[%CLASS_GREEN%]","");
  htmlOutput = htmlOutput.replace("[%CLASS_RED%]","");
  
  return htmlOutput;
  
}

function runReportsLoader() {
  var cache = CacheService.getDocumentCache();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss_sheet = ss.getSheetByName(report_configuration_name);
  cache.put(html_file_name_ref, "modal_report_run_summary", 1500);
  modal_title = "Report Status";
  modal_html_name = "loader";
  modal_width = 400;
  modal_height = 150;
  if (ss_sheet == null) {
    cache.put(html_file_name_ref, "modal_report_status_ops", 1500);
    showModalDialogCustom(modal_width,modal_height,modal_html_name,modal_title);
    return true;
  }
  modal_height = 270;
  showModalDialogCustom(modal_width,modal_height,modal_html_name,modal_title);
}

function generaReportConfiguration(form) {
  var cache = CacheService.getDocumentCache();
  
  var report_name = form.report_name;
  var select_account = form.select_account;
  var select_property = form.select_property;
  var select_view = form.select_view;
  var report_m = form.report_m;
  var report_d = form.report_d;
  var report_s = form.report_s;

  //**********
  // Styling Report Configuration
  //**********
  var arr_fields_len = arr_fields.length;
  var last_row_after_fields = arr_fields_len + 2;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report_configuration_sheet = createNewSheetByName(ss, report_configuration_name);
  var sh = ss.getSheetByName(report_configuration_name);
  
  if (report_configuration_sheet == '0') {
    // PRIMA COLONNA
    sh.getRange('A1').setBackground('#757575').setValue('Configuration Options').setFontWeight('bold').setFontSize(12).setFontColor("#FFFFFF").setVerticalAlignment('middle');
    sh.setColumnWidth(1, 240).setRowHeight(1, 40);
    sh.getRange(2,1,arr_fields_len,1).setBackground('#efefef').setValues(arr_fields).setFontSize(10).setFontWeight('bold');
    sh.setRowHeights(2, arr_fields_len, 24);
    sh.getRange((last_row_after_fields+1),1,(1000-last_row_after_fields+1),1).setBackground('#efefef').mergeVertically(); // per escludere dal merge la riga suill'informazione dei campi con * obbligatori
    sh.setFrozenColumns(1);
    // PRIMA RIGA DALLA SECONDA COLONNA
    sh.getRange('B1:Z1').mergeAcross().setBackground('#ed750a').setValue('Your Google Analytics Unsampled Reports').setFontWeight('bold').setFontSize(12).setFontColor("#FFFFFF").setVerticalAlignment('middle');
    sh.getRange('B2:Z2').setBackground('#F5F5F5').setFontColor("#444444").setFontWeight('bold').setHorizontalAlignment('left');
    sh.setColumnWidths(2, 25, 180);
    // RIGHE CENTRALI
    sh.getRange('B2:Z'+arr_fields_len+1).setFontSize(10).setFontColor("#444444").setHorizontalAlignment('left');
    sh.getRange('B2:Z'+arr_fields_len+1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sh.getRange('B2:Z'+arr_fields_len+1).setVerticalAlignment("bottom");
    // PENULTIME 6 RIGHE (CREATED, UPDATED, STATUS, REPORT ID, DOCUMENT ID, NEXT SCHEDULED CHECK )
    sh.getRange('A'+(arr_fields_len-4)+':Z'+(arr_fields_len-4)).setBorder(true, null, null, null, false, true, "#999999", null);
    sh.getRange('A'+(arr_fields_len-4)+':Z'+(arr_fields_len+1)).setFontStyle("italic")
    sh.getRange('B'+(arr_fields_len-4)+':Z'+(arr_fields_len+1)).setBackground('#F5F5F5').setFontWeight('normal');
    // ULTIMA RIGA DALLA PRIMA COLONNA
    sh.getRange(last_row_after_fields,1,1,1).setValue('Fields with * are mandatory').setFontSize(9).setFontColor("#444444").setBackground('#efefef').setHorizontalAlignment('left').setVerticalAlignment('middle');
    // ULTIMA RIGA DALLA SECONDA COLONNA
    sh.getRange(last_row_after_fields,2,1,1).setValue('For help with this add-on:').setFontSize(9).setFontColor("#444444").setHorizontalAlignment('right').setVerticalAlignment('middle');
    sh.getRange('C'+last_row_after_fields+':Z'+last_row_after_fields).mergeAcross().setValue('https://www.analyticstraps.com/google-analytics-unsampled-reports-spreadsheet-add-on/').setFontSize(9).setFontColor("#1155cc").setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeights(last_row_after_fields, 1, 28);
    var white_range = sh.getRange(last_row_after_fields+1,2,1000-last_row_after_fields+1,25).merge();
    sh.setActiveRange(white_range);
    SpreadsheetApp.flush();
  }
  
  // Condizioni per verificare quale è la prima colonna libera dove inserire le opzioni del report in quando le prime due colonne verrebbero saltate da getLastColumn() per via delle righe con il link del sito
  var bln_lastColumn_detected = true;
  var lastColumn = 1;
  var column_option_range_detect_free = sh.getRange('B2:B'+(arr_fields_len+1)).getValues();
  var column_option_range_detect_free_len = column_option_range_detect_free.length;
  for (var i=0; i<column_option_range_detect_free_len;i++) {
    if (column_option_range_detect_free[i] != '') {
        bln_lastColumn_detected = false;
    }
  }
  if (!bln_lastColumn_detected) {
    bln_lastColumn_detected = true;
    lastColumn = 2;
    column_option_range_detect_free = sh.getRange('C2:C'+(arr_fields_len+1)).getValues();
    column_option_range_detect_free_len = column_option_range_detect_free.length;
    for (var i=0; i<column_option_range_detect_free_len;i++) {
      if (column_option_range_detect_free[i] != '') {
        bln_lastColumn_detected = false;
      }
    }
  }
  if (!bln_lastColumn_detected) {
    lastColumn = sh.getLastColumn();
  }
var arr_fields_values = [[report_name], [select_account], [select_property], [select_view], ['30daysAgo'], ['yesterday'], [report_m], [report_d], [''], [report_s], [''], [''], [''], [''], [''], ['']];
  var arr_fields_values_len = arr_fields_values.length;
  sh.getRange(2,lastColumn+1,arr_fields_values_len,1).setValues(arr_fields_values);

  return true;
}

function getSegments() {
  var arr_segments_list = [];
  var segments_list = Analytics.Management.Segments.list();
  var total_results = segments_list.totalResults;
  for (var s=0; s < total_results; s++) {
      segments_segmentId = segments_list.items[s].segmentId;
      segments_name = segments_list.items[s].name;
      segments_type = segments_list.items[s].type; // BUILT_IN, CUSTOM
      segments_definition = segments_list.items[s].definition;
      arr_segments_list.push([segments_type, segments_segmentId, segments_name, segments_definition]);
  }
  return arr_segments_list;
}

function getMetadata() {
  var metadata_status_required = "PUBLIC"; // PUBLIC, DEPRECATED
  var arr_metadata_list = [];
  var metadata_list = Analytics.Metadata.Columns.list('ga');
  var total_results = metadata_list.totalResults;
  for (var md=0; md < total_results; md++) {
    metadata_status = metadata_list.items[md].attributes["status"];
    if (metadata_status == metadata_status_required) {
      metadata_group = metadata_list.items[md].attributes["group"];
      metadata_type = metadata_list.items[md].attributes["type"];
      metadata_id = metadata_list.items[md].id;
      metadata_uiname = metadata_list.items[md].attributes["uiName"];
      arr_metadata_list.push([metadata_group, metadata_type, metadata_id, metadata_uiname]);
    }
  }
  return arr_metadata_list;
}

function get360() {
  var web_properties_level_required = "PREMIUM"; // STANDARD, PREMIUM
  var arr_accounts_properties_views_360 = [];
  var premium_counter = 0;
  var web_properties_len = 0;
  var web_profiles_len = 0;
  var web_property_level = '';
  var account_summaries_list = Analytics.Management.AccountSummaries.list();
  var total_results = account_summaries_list.totalResults;
  for (var i=0; i < total_results; i++) {
    premium_counter = 0;
    web_properties_len = account_summaries_list.items[i].webProperties.length;
    for (var j=0; j<web_properties_len; j++) {
      web_property_level = account_summaries_list.items[i].webProperties[j].level;
      if (web_property_level == web_properties_level_required) {
        web_profiles_len = account_summaries_list.items[i].webProperties[j].profiles.length;
        for (var k=0; k<web_profiles_len; k++) {
          arr_accounts_properties_views_360.push([[account_summaries_list.items[i].id], [account_summaries_list.items[i].name], [account_summaries_list.items[i].webProperties[j].id], [account_summaries_list.items[i].webProperties[j].name], [account_summaries_list.items[i].webProperties[j].profiles[k].id], [account_summaries_list.items[i].webProperties[j].profiles[k].name]]);
        }
      }
    }
  }
  // only for test
  //arr_accounts_properties_views_360.push([['160739218'], ['Test Account'], ['UA-160739218-1'], ['Test Property'], ['213646178'], ['Test View']]);
  return arr_accounts_properties_views_360;
}

// Crea un nuovo Sheet in base al nome se non è già presente
function createNewSheetByName(ss, sn) {
  var ss_sheet = ss.getSheetByName(sn);
  if (ss_sheet != null) { return '1'; }
  ss.insertSheet().setName(sn);
  return '0';
}

function showModalDialogCustom(wid,hig,htm_n,ttl) {
  var htmlOutput = HtmlService
  .createHtmlOutputFromFile(htm_n)
  .setWidth(wid)
  .setHeight(hig);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ttl);
}

function browserMsgBox(msg) {
  Browser.msgBox(msg);
}

// Effettua la richiesta di report non campionato
function insert_UnsampledReport(resource, select_account, select_property, select_view) {
    try{
      return Analytics.Management.UnsampledReports.insert(resource, select_account, select_property, select_view);
    } catch(e) {
      Logger.log(e.message);
      Browser.msgBox(e.message);
    };
  //return Analytics.Management.UnsampledReports.insert(resource, select_account, select_property, select_view);
}

// Recupera un report non campionato sulla base del suo ID
function get_UnsampledReport(select_account, select_property, select_view, report_id) {
  return Analytics.Management.UnsampledReports.get(select_account, select_property, select_view, report_id);
}

function importReportFromDrive(ss, title, fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var csvString = file.getBlob().getDataAsString();
    var data = Utilities.parseCsv(csvString);
    } catch(e) {
      Logger.log(e.message);
      Browser.msgBox(e.message);
    };

  var sheet = ss.getSheetByName(title);
  if (sheet == null) {
    sheet = ss.insertSheet().setName(title);
  }

  sheet.clear();
  var range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
  return true;
}

// Crea un trigger in modo programmatico basato sul tempo (e lo inserisce nelle Properies)
function createTriggerBasedTime(prop, script_trigger_n, func_name_to_call, dateObj) {
  var triggers = ScriptApp.getProjectTriggers();
  var trigger_timebased_runtime = ScriptApp.newTrigger(func_name_to_call).timeBased().at(new Date(dateObj)).inTimezone(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()).create();
  prop.setProperty(script_trigger_n, trigger_timebased_runtime.getUniqueId());
}

function millisToMinutesAndSeconds(millis, withseconds) {
  var minutes = Math.floor(millis / 60000);
  var seconds = ((millis % 60000) / 1000).toFixed(0);
  if (withseconds) {
    return minutes + ":" + (seconds < 10 ? '0' : '') + seconds;
  }
  return minutes;
}


// Elimina tutte le properties
/*
function deleteProps() {
  var docProps = PropertiesService.getDocumentProperties();
  docProps.deleteAllProperties()
}
*/

// Verifica i triggers attualmente esistenti (solo per TEST)
/*
function testTriggers() {
  var docProps = PropertiesService.getDocumentProperties();
  var defaultTrigger_id = docProps.getProperty(document_trigger_properties_name);
  Logger.log('id: ' + defaultTrigger_id);
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger){
    try {
      if (trigger.getUniqueId() != defaultTrigger_id) {
        //ScriptApp.deleteTrigger(trigger);
        Logger.log('dentro: ' + trigger.getUniqueId());
      }
      Logger.log('fuori: ' + trigger.getUniqueId());
    } catch(e) {
      throw e.message;
    };
    Utilities.sleep(500);
  });
};
*/

// Rimuove un trigger in base al suo id recuperato dalle Properties (utilizzato per rimuovere i trigger temporizzati ad un orario specifico inseriti in modo programmatico)
// non più utilizzato
/*
function deleteTriggerById(prop,script_trigger_n) {
  var trigger_timebased_runtime_id = prop.getProperty(script_trigger_n); 
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (trigger_timebased_runtime_id != null) {
      if (triggers[i].getUniqueId() == trigger_timebased_runtime_id) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  }  
}
*/

// Rimuove tutti i trigger del progetto
// non più utilizzato
/*
function deleteTriggers(docProps) {
  var defaultTrigger_id = docProps.getProperty(document_trigger_properties_name); 
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger){
    try {
      if (trigger.getUniqueId() != defaultTrigger_id) {
        ScriptApp.deleteTrigger(trigger);
      }
    } catch(e) {
      throw e.message;
    };
    Utilities.sleep(500);
  });
};
*/

// Elimina tutti i triggers e le properties (solo per TEST o la prima volta che si usa l'add-on o se rimangono trigger appesi dopo che il foglio principale del report viene cancellato)
function delete_ALLTriggers_ALLProperties() {
  var docProps = PropertiesService.getDocumentProperties();
  docProps.deleteAllProperties()
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger){
    try {
      ScriptApp.deleteTrigger(trigger);
    } catch(e) {
      throw e.message;
    };
    Utilities.sleep(500);
  });
};

function setDefaultTrigger(docProps) {
  var defaultTrigger = ScriptApp.newTrigger('runReportsHourly').timeBased().everyHours(1).create();
  docProps.setProperty(document_trigger_properties_name, defaultTrigger.getUniqueId());
}

function runReportsHourly() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss_sheet = ss.getSheetByName(report_configuration_name);

  if (!ss_sheet) { // se il foglio principale non esiste perché ad esempio è stato cancellato
    delete_ALLTriggers_ALLProperties();
    return;
  }
  
  var documentProperties = PropertiesService.getDocumentProperties();
  
  // Verifico la presenza di una schedulazione salvata e se le informazioni corrispondono al momento della scansione oraria reinizializzo i report svuotando le celle con i valori di controllo
  var scheduled_setup = documentProperties.getProperty(document_schedule_properties_name);  
  var blnActivateSchedule = false;
  if (scheduled_setup) {

    //var d = new Date();
    var dt = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    var d_yr = dt.substring(0, 4);
    var d_mh = dt.substring(5, 7);
    var d_dy = dt.substring(8, 10);
    var d_hr = dt.substring(11, 13);
    var d_mt = dt.substring(14, 16);
    var d_sd = dt.substring(17, 19);

    var d = new Date(d_yr, parseInt(d_mh-1), d_dy, d_hr, d_mt, d_sd);

    var arr_scheduled_setup = scheduled_setup.split("|");
    var interval = arr_scheduled_setup[0];
    var dayOfWeek = arr_scheduled_setup[1];
    var dayOfMonth = arr_scheduled_setup[2];
    var hourOfDay = arr_scheduled_setup[3];
    switch(interval) {
      case '1':
        // Esempio: 1|0|0|19 significa ogni giorno alle 19:00
        if (d.getHours() == hourOfDay) {
          blnActivateSchedule = true;
        }
        break;
      case '2':
        dayOfWeek = (parseInt(dayOfWeek) + 1);
        if (d.getDay() == dayOfWeek) {
          if (d.getHours() == hourOfDay) {
            blnActivateSchedule = true;
          }
        }
        break;
      case '3':
        dayOfMonth = (parseInt(dayOfMonth) + 1);
        if (d.getDate() == dayOfMonth) {
          if (d.getHours() == hourOfDay) {
            blnActivateSchedule = true;
          }
        }
        break;
      default:
    }
  }
  if (blnActivateSchedule) {
    ss_sheet.getRange("B12:Z17").clearContent();
    documentProperties.setProperty(document_autorerun_properties_name, '1'); // valorizzo la property per indicare preventivamente la necessità di effettuare il run automatico per i report pending (successivamente se non ci sono report pending la svuoto, in questo modo il rerun automatico non entra più in questa funzione)
  }

  var document_autorerun_properties_name_value = documentProperties.getProperty(document_autorerun_properties_name);

  if (document_autorerun_properties_name_value) { runReports(); } // se la proprietà legata alla presenza di report pending è valorizzata allora esegui la funzione. Nota: può essere true perché il report è stato avviato manualmente o perché ci sono le condizioni dello schedule
  
}

function scheduleReportByTrigger(form) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var document_trigger_properties_name_value = documentProperties.getProperty(document_trigger_properties_name);
  if (!document_trigger_properties_name_value) { // controllo se l'id del trigger di default esiste, altrimenti creo il trigger orario (qui serve nel caso in cui l'utente inserisca lo schedule senza aver mai eseguito manualmente il report. Non è inserito in runReports poiché lì, se non esiste il trigger di default, significa che ci si trova a situazione iniziale o con l'add-on vecchio per cui reinizializzo tutto e verrebbe eliminato il trigger dello schedule)
    delete_ALLTriggers_ALLProperties();
    setDefaultTrigger(documentProperties);
  }
  if (!Array.isArray(form.automate)) {
    documentProperties.deleteProperty(document_schedule_properties_name);
    //deleteTriggers(documentProperties); // elimina tutti i triggers (ovvero solo quello schedulato e fa pulizia a chi ha installato il vecchio add-on) ad eccezione di quello di default
    //setDefaultSchedule(documentProperties);
  } else {
    //deleteTriggers(documentProperties); // elimina tutti i triggers (ovvero solo quello schedulato e fa pulizia a chi ha installato il vecchio add-on) ad eccezione di quello di default
    //setDefaultSchedule(documentProperties);
    var interval = form.interval;
    var dayOfWeek = parseInt(form.dayOfWeek);
    var dayOfMonth = parseInt(form.dayOfMonth);
    var hourOfDay = parseInt(form.hourOfDay);

    documentProperties.setProperty(document_schedule_properties_name, interval+'|'+dayOfWeek+'|'+dayOfMonth+'|'+hourOfDay); // salvo la selezione dello schedule nelle Properties

// rimosso il codice seguente dalla verisone 1.1.0 poiché l'errore nelle add-on "Questo componente aggiuntivo ha creato troppi trigger time-based nel documento per questo account utente Google" non significa che non possono essere creati più di 20 triggers per documento per utente bensì che non possono essere creati più di un trigger dello stesso tipo */
/*
    switch(interval) {
      case '1':
        ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().everyDays(1).atHour(hourOfDay).create();
        break;
      case '2':
        if (dayOfWeek == '0') {
          ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(hourOfDay).create();
        } else if (dayOfWeek == '1') {
          ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(hourOfDay).create();
        } else if (dayOfWeek == '2') {
          ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(hourOfDay).create();
        } else if (dayOfWeek == '3') {
          ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(hourOfDay).create();
        } else if (dayOfWeek == '4') {
          ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(hourOfDay).create();
        } else if (dayOfWeek == '5') {
          ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(hourOfDay).create();
        } else {
          ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(hourOfDay).create();
        }
        break;
      case '3':
        dayOfMonth = dayOfMonth + 1;
        ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().onMonthDay(1).atHour(hourOfDay).create();
        break;
      default:
        //ScriptApp.newTrigger('clearContentForNewRequestAndRunReport').timeBased().everyHours(1).create(); // il valore 0 non è più selezionabile poiché l'intervallo ogni ora è presente di default
    }
*/
  }
}

/*
function setPropertyForTest() {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(document_schedule_properties_name, '1|0|0|22');
}
*/

/*
function simulateTriggerForTest() { // Funzionalità integrata nella funzione runReports()

// return_date = Utilities.formatDate(new Date(date_to_check), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");

  var documentProperties = PropertiesService.getDocumentProperties();
  var scheduled_setup = documentProperties.getProperty(document_schedule_properties_name);
  
  var blnActivateSchedule = false;
  if (scheduled_setup) {

    //var d = new Date();
    var dt = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    var d_yr = dt.substring(0, 4);
    var d_mh = dt.substring(5, 7);
    var d_dy = dt.substring(8, 10);
    var d_hr = dt.substring(11, 13);
    var d_mt = dt.substring(14, 16);
    var d_sd = dt.substring(17, 19);

    var d = new Date(d_yr, parseInt(d_mh-1), d_dy, d_hr, d_mt, d_sd);
    
    var arr_scheduled_setup = scheduled_setup.split("|");
    var interval = arr_scheduled_setup[0];
    var dayOfWeek = arr_scheduled_setup[1];
    var dayOfMonth = arr_scheduled_setup[2];
    var hourOfDay = arr_scheduled_setup[3];
      
    switch(interval) {
      case '1':
        // Esempio: 1|0|0|19 significa ogni giorno alle 19:00
        if (d.getHours() == hourOfDay) {
          blnActivateSchedule = true;
        }
        break;
      case '2':
        dayOfWeek = (parseInt(dayOfWeek) + 1);
        if (d.getDay() == dayOfWeek) {
          if (d.getHours() == hourOfDay) {
            blnActivateSchedule = true;
          }
        }
        break;
      case '3':
        dayOfMonth = (parseInt(dayOfMonth) + 1);
        if (d.getDate() == dayOfMonth) {
          if (d.getHours() == hourOfDay) {
            blnActivateSchedule = true;
          }
        }
        break;
      default:
        
    }
 
  }
  
  Browser.msgBox(dt + ' - ' + d + ' | ' + blnActivateSchedule + ' - ' + arr_scheduled_setup);
  Browser.msgBox(d.getHours() + ' - ' + hourOfDay)
  Browser.msgBox(d.getDay() + ' - ' + (parseInt(dayOfWeek) + 1))
  Browser.msgBox(d.getDate() + ' - ' + (parseInt(dayOfMonth) + 1))
  
  //if (blnActivateSchedule) {
  //  ss_sheet.getRange("B12:Z16").clearContent();
  //}
}
*/

function getScheduleTriggerDetails() {
  var documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperty(document_schedule_properties_name); 
}

/*
// non più utilizzata (era la funzione attivata dai trigger schedulati ma a causa dell'impossibilità nelle add-on di installare 2 trigger dello stesso tipo ho utilizzato un workaround che sfrutta il trigger orario e verifica le impostazioni di schedulazione per rinizializzare i report)
function clearContentForNewRequestAndRunReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ss_sheet = ss.getSheetByName(report_configuration_name);
  if (!ss_sheet) { // se il foglio principale non esiste perché ad esempio è stato cancellato
    delete_ALLTriggers_ALLProperties();
    return;
  }
  ss_sheet.getRange("B12:Z16").clearContent(); // svuoto i campi solo per il report schedulato in questo modo tolgo il COMPLETED e viene rieffettuata la richiesta ex-novo (ad esempio perché la data =TODAY()-1 è variata)
  runReports();
}
*/

function validateReportDate(date_to_check) {
  var return_date = date_to_check;
  if (isValidDate(new Date(date_to_check))) {
    return_date = Utilities.formatDate(new Date(date_to_check), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
  }
  var date_regex = new RegExp("[0-9]{4}-[0-9]{2}-[0-9]{2}|today|yesterday|[0-9]+(daysAgo)");
  var date_checked = (return_date+"").match(date_regex);
  if (!date_checked) { return_date = "";  }
  return return_date;
}

function isValidDate(d) {
  return d instanceof Date && !isNaN(d);
}
