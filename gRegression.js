// This one is required to make regression-js work, as it expects a window variable.
// Thanks to Tom Alexander for the library. See source at
// https://github.com/Tom-Alexander/regression-js
if (typeof window == 'undefined') {
  var window = {};
}

// Add menu to the spreadsheet.
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [];
  entries.push({name : 'Run regression analysis', functionName : 'regressionDialog'});
  entries.push({name : 'Help', functionName : 'help'});
  sheet.addMenu('gRegression', entries);
};
function onInstall() {
  onOpen();
}

// The dialog for the regression analysis settings.
function regressionDialog() {
  var app = UiApp.createApplication().setTitle('Google spreadsheet regression analysis 1.0');
  var handler = app.createServerHandler('regressionHandler');

  // Build a list of regression modes.
  var modes = {
    linear : 'Linear function ax + b',
    exponential : 'Exponential function ae^bx',
    logarithmic : 'Logarithmic function a + bln(x)',
    power : 'Power function ax^b',
    polynomial : 'Polynomial function a0x^0 ... + anx^n',
  };
  app.add(app.createLabel('Select type of function'));
  var modeList = app.createListBox().setName('mode')
  for (var mode in modes) {
    modeList.addItem(modes[mode], mode);
  }
  app.add(modeList);
  handler.addCallbackElement(modeList);

  // Add some extras used only for polynomial functions.
  var visibilitySwitcher = app.createServerHandler('visibilitySwitcher').addCallbackElement(modeList);
  modeList.addChangeHandler(visibilitySwitcher);
  app.add(app.createLabel('…of order').setId('polynomialLabel').setVisible(false));
  var order = app.createListBox().setName('polynomialOrder').setId('polynomialOrder').setVisible(false).addItem('2', 2).addItem('3', 3).addItem('4', 4).addItem('5', 5);
  handler.addCallbackElement(order);
  app.add(order);

  // Some extra options.
  app.add(app.createHTML('<br />'));
  var addSheet = app.createCheckBox('Create a new sheet with regression-fitted data and a scatter plot') .setName('addSheet');
  app.add(addSheet);
  handler.addCallbackElement(addSheet);

  // Add buttons for ok an cancel.
  app.add(app.createHTML('<hr />'));
  var close = app.createServerHandler('close');
  var cancel = app.createButton('Cancel').setId('cancel').addClickHandler(close);
  app.add(cancel);
  var ok = app.createButton('Run').setId('ok').addClickHandler(handler);
  app.add(ok);

  // Run some validations.
  var errorMessage = app.createLabel().setId('errorMessage');
  app.add(errorMessage);
  if (SpreadsheetApp.getActiveRange().getNumColumns() != 2) {
    errorMessage.setText('Error: You must select two columns, containing x and y values.');
    ok.setEnabled(false);
  }
  if (SpreadsheetApp.getActiveRange().getNumRows() < 2) {
    errorMessage.setText('Error: You must select at least two rows, containing coordinates for points.');
    ok.setEnabled(false);
  }

  // Add some cred.
  app.add(app.createLabel('Thanks to Tom Alexander for the regression analysis code. See https://github.com/Tom-Alexander/regression-js for details.'));

  SpreadsheetApp.getActiveSpreadsheet().show(app);
}

// Turn on/off visibility for polyonmial order.
function visibilitySwitcher(eventInfo) {
  var app = UiApp.getActiveApplication();
  var label = app.getElementById('polynomialLabel');
  var order = app.getElementById('polynomialOrder');
  if (eventInfo.parameter.mode == 'polynomial') {
    label.setVisible(true);
    order.setVisible(true);
  }
  else {
    label.setVisible(false);
    order.setVisible(false);
  }
  return app;
}

// The function coordinating the regression analysis.
function regressionHandler(eventInfo) {
  if (eventInfo.parameter.mode == 'polynomial' && SpreadsheetApp.getActiveRange().getNumRows() <= parseInt(eventInfo.parameter.polynomialOrder)) {
    Browser.msgBox('Error', 'You need more data points to do polynomial regression of this order. Select more rows and try again.', null);
    return UiApp.getActiveApplication().close();
  }

  // Get the best fit, and display it.
  var values = SpreadsheetApp.getActiveRange().getValues();
  var result = window.regression(eventInfo.parameter.mode, values, parseInt(eventInfo.parameter.polynomialOrder));
  Browser.msgBox('Best fit', result.string, null);

  // Build a new sheet, if asked to.
  if (eventInfo.parameter.addSheet == 'true') {
    var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    newSheet.activate();
    // Setting some titles.
    newSheet.getRange('A1').setValue('Function:');
    newSheet.getRange('B1').setValue(result.string);
    newSheet.getRange('A2').setValue('x');
    newSheet.getRange('B2').setValue('y (original)');
    newSheet.getRange('C2').setValue('y (fitted)');

    // Fill in the values from the original selection, and the regression.
    newSheet.getRange(3, 1, values.length, 2).setValues(values);
    var row = 3;
    for (var i in result.points) {
      newSheet.getRange(row, 3).setValue(result.points[i][1]);
      row++;
    }

    // Build a scatter chart.
    var dataRange = newSheet.getRange(2, 1, values.length, 3);
    var chart = newSheet.newChart()
      .setChartType(Charts.ChartType.SCATTER)
      .addRange(dataRange)
      .setPosition(1, 4, 0, 0)
      .setOption('title', result.string)
      .build();
    newSheet.insertChart(chart);
  }
}

// Simple handler for closing the active UI.
function close(eventInfo) {
  return UiApp.getActiveApplication().close();
}

// Displays help information.
function help() {
  var app = UiApp.createApplication().setTitle('Google spreadsheet regression analysis 1.0');
  var handler = app.createServerHandler('close');
  app.add(app.createHTML('<strong>How to use:</strong> Select x and y data by click-and-drag in the spreadsheet, then run "regression analysis" from the menu and select the relevant options. (Data must be in columns, with x values in the first selected column.)<br/>'));
  app.add(app.createHTML('Use the option "Create a new sheet with regression-fitted data and a scatter plot" to have the function-fitted data plotted in a chart with no extra work.<br/><br/>'));
  app.add(app.createHTML('Thanks to Thanks to Tom Alexander for the JavaScript library for regression analysis: https://github.com/Tom-Alexander/regression-js)<br/><br/>'));
  app.add(app.createHTML('Follow the link below to find more information about this project, including source code and issue queue.<br/><br/>'));
  app.add(app.createAnchor('gRegression project page on GitHub', 'https://github.com/Itangalo/gRegression'));
  app.add(app.createButton('OK!', handler));
  SpreadsheetApp.getActiveSpreadsheet().show(app);
}

function debug(value) {
  SpreadsheetApp.getActiveSpreadsheet().toast(value, typeof value);
}
