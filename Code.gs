function onOpen() {
    //  var response = Browser.msgBox('Update?', Browser.Buttons.YES_NO);
    //  
    //  if (response == "yes"){
    //    readStudentEvents();
    //  }
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('AutoMenu')
        .addItem('Update', 'readStudentEvents')
        .addToUi();
}

function doGet(e) {

    return HtmlService.createHtmlOutputFromFile('Index')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);


}

function migrate() {

    var calendars = CalendarApp.getAllCalendars();
	Logger.log(calendar);

}

function getData() {

    readStudentEvents();
    var spread = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spread.getSheets();

    var monthSheets = []

    for (var i = spread.getNumSheets(); i--; i >= 0) {
        name = sheets[i].getName();
        if (name.indexOf("m_") != -1) {
            monthSheets = monthSheets.concat(name);
        }
    }

    Logger.log(monthSheets);

    sheet_one = spread.getSheetByName(monthSheets[0]);
    sheet_zero = spread.getSheetByName(monthSheets[1]);
    sheet_minus_one = spread.getSheetByName(monthSheets[2]);

    value_one = sheet_one.getSheetValues(2, 7, 1, 1);
    value_zero = sheet_zero.getSheetValues(2, 7, 1, 1);
    value_minus_one = sheet_minus_one.getSheetValues(2, 7, 1, 1);

    var data = [value_one.toString(), value_zero.toString(), value_minus_one.toString()];

    //  Logger.log(data);
    return data;

}

function readStudentEvents() {

    //  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    //     .alert('You clicked the update button');

    var spread = SpreadsheetApp.getActiveSpreadsheet();
    var first_sheet = spread.getSheetByName("Main");
    var fees_range = first_sheet.getSheetValues(1, 1, first_sheet.getLastRow(), 2);

    var sheets = spread.getSheets();

    var monthSheets = []

    for (var i = spread.getNumSheets(); i--; i >= 0) {
        name = sheets[i].getName();
        if (name.indexOf("m_") != -1) {
            monthSheets = monthSheets.concat(name);
        }
    }

    update_sheets = checkMonths(monthSheets);
    calendars = getCalendars();
    Logger.log(update_sheets);

    var sheet_content = [];

    for (var i = 0; i < update_sheets.length; i++) {
        sheet_content.push({
            sheet_name: update_sheets[i],
            teacher: []
        });

        var sheet_name_array = update_sheets[i].split(/-|_/);
        var start_date = new Date(parseInt(sheet_name_array[2]), parseInt(sheet_name_array[1]) - 1, 1, 0, 0, 0, 0);
        var end_date = new Date();
        end_date.setTime(start_date.getTime());
        end_date.setMonth(end_date.getMonth() + 1);

        for (var j = 0; j < calendars.length; j++) {
            sheet_content[i].teacher.push({
                name: calendars[j],
                events: []
            });

            var cal = CalendarApp.getCalendarsByName(calendars[j]);
            events = cal[0].getEvents(start_date, end_date);
            for (k = 0; k < events.length; k++) {
                if (events[k].getLocation() == "CEAP") {
                    sheet_content[i].teacher[j].events.push(events[k]);
                }
            }
            Logger.log(start_date + "   " + end_date + "   " + update_sheets[i] + "   " + events.length);

        }
    }
    log_events(sheet_content);
}

function checkMonths(months) {
    var sheet_names = [];

    var utcDate2 = new Date();
    var c_month_start = new Date(Date.UTC(utcDate2.getUTCFullYear(), utcDate2.getUTCMonth(), 1, 0, 0, 0));
    Logger.log("c_month " + c_month_start);

    var next_month = new Date();
    next_month.setTime(c_month_start.getTime());
    next_month.setMonth(next_month.getMonth() + 1);
    next_month.setHours(next_month.getHours() + 1); //para lidar com diferenças de daylight savings
    Logger.log("n_month " + next_month);

    var previous_month = new Date();
    previous_month.setTime(c_month_start.getTime());
    previous_month.setMonth(previous_month.getMonth() - 1);
    previous_month.setHours(previous_month.getHours() - 1); //para lidar com diferenças de daylight savings
    Logger.log("p_month " + previous_month);

    for (var i = months.length; i--; i >= 0) {
        var month_array = months[i].split(/-|_/);

        //Criterio para ação: Mês presente, anterior ou seguinte
        var month_start = new Date(Date.UTC(parseInt(month_array[2]), parseInt(month_array[1]) - 1, 1, 0, 0, 0));
        Logger.log(month_start);

        if (month_start.getTime() <= next_month.getTime() && month_start.getTime() >= previous_month.getTime()) {
            sheet_names = sheet_names.concat(months[i]);
        }

        Logger.log(months[i]);
    }
    Logger.log(sheet_names);

    return sheet_names;
}

function getCalendars() {
    var teacher_calendars = [];

    var calendars = CalendarApp.getAllCalendars();

    for (var i = calendars.length; i--; i >= 0) {
        var calendar_name = calendars[i].getName();

        if (calendar_name.indexOf("joao") != -1) {
            teacher_calendars = teacher_calendars.concat(calendar_name);
        }
    }

    return teacher_calendars;
}

function log_events(sheet_content) {
    spread_var = SpreadsheetApp.getActiveSpreadsheet();
    for (sheet in sheet_content) {
        Logger.log(sheet_content[sheet].sheet_name);
        sheet_var = spread_var.getSheetByName(sheet_content[sheet].sheet_name);
        sheet_var.clear();
        header = sheet_var.getRange(1, 1, 1, 7);
        header.setValues([
            ["Nome", "Tipologia", "Total (h)", "Rate EE", "Rate Prof.", "Total EE", "Total Prof."]
        ]);

        //putting together array for totals table
        var teacher_totals_array = [
            ["Teacher Name", "Teacher Name", "Total a Receber"]
        ];
        ///

        for (teacher in sheet_content[sheet].teacher) {
            Logger.log(sheet_content[sheet].teacher[teacher].name);
            last = sheet_var.getLastRow();
            name_range = sheet_var.getRange(last + 1, 1);

            name_range.setValue(sheet_content[sheet].teacher[teacher].name);
            student_sum = [];

            for (event in sheet_content[sheet].teacher[teacher].events) {
                calendar_event = sheet_content[sheet].teacher[teacher].events[event];
                event_title_array = calendar_event.getTitle().split(" ");
                duration = (calendar_event.getEndTime() - calendar_event.getStartTime()) / (1000 * 60 * 60);
                type = event_title_array[0];
                student_name = event_title_array.slice(1).join(" ");

                stored = 0;
                for (i in student_sum) {
                    if (student_sum[i][0].indexOf(student_name) != -1 && student_sum[i][1].indexOf(type) != -1) {
                        student_sum[i][2] = student_sum[i][2] + duration;
                        stored = 1;
                        break
                    }
                }
                if (stored == 0) {
                    student_sum.push([student_name, type, duration]);
                }
            }

            //piece for calculating teacher totals
            ee_sum_range = sheet_var.getRange(last + 1, 6);
            row_range = [ee_sum_range.getRow() + 1, ee_sum_range.getRow() + student_sum.length];
            teacher_sum_range = sheet_var.getRange(last + 1, 7);

            if (student_sum == 0) {
                ee_sum_range.setValue(0);
                teacher_sum_range.setValue(0);
            } else {
                ee_sum_range.setValue("=SUM(F" + row_range[0] + ":F" + row_range[1] + ")");
                teacher_sum_range.setValue("=SUM(G" + row_range[0] + ":G" + row_range[1] + ")");
            }
            ///

            Logger.log(student_sum);
            extended = add_rates(student_sum);

            for (line in extended) {
                last = sheet_var.getLastRow();
                name_range = sheet_var.getRange(last + 1, 1, 1, 7);
                //Logger.log(name_range.getA1Notation());
                row = last + 1;
                name_range.setValues([extended[line].concat("=C" + row + "*D" + row, "=C" + row + "*E" + row)]);
            }

            teacher_totals_array = teacher_totals_array.concat([
                [sheet_content[sheet].teacher[teacher].name, ee_sum_range.getValue().toString(), teacher_sum_range.getValue().toString()]
            ]);
            //Logger.log(teacher_totals_array);
            //Logger.log(extended);
        }

        /////print teacher's final income
        ///uncomment for-loop for condensing informations with multiple teachers
        //  for (i in teacher_totals_array){
        //    var row = +i + +1;
        //    var teacher_totals_range = sheet_var.getRange(row, 9, 1, 3);
        //    Logger.log([teacher_totals_array[i]]);
        //    Logger.log(teacher_totals_range.getA1Notation());
        //    teacher_totals_range.setValues([teacher_totals_array[i]])
        //    
        //    }
        /////////////////////
    }
}

function add_rates(sum_array) {

    var spread = SpreadsheetApp.getActiveSpreadsheet();
    var first_sheet = spread.getSheetByName("Main");
    var fees_range = first_sheet.getSheetValues(1, 1, first_sheet.getLastRow(), 5);
    //add get type ratings (parents and teacher) to push and functions to compute totals
    //Logger.log(fees_range);
    extended = [];

    for (i in sum_array) {
        name = sum_array[i][0];
        type = sum_array[i][1];

        for (var j = 1; j < fees_range.length; j++) {
            if (fees_range[j][0].indexOf(name) != -1) {
                if (type == "Exp.") {
                    rate_parent = fees_range[j][1];
                    rate_teacher = fees_range[j][3];
                    break
                } else if (type == "EA") {
                    rate_parent = fees_range[j][2];
                    rate_teacher = fees_range[j][4];
                    break
                } else {
                    rate_parent = 0;
                    rate_teacher = 0;
                }
            }
        }
        extended.push(sum_array[i].concat(rate_parent, rate_teacher))
        rate_parent = 0;
        rate_teacher = 0;
    }
    Logger.log(extended);
    return extended;
}