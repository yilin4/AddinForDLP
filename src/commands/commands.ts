/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

async function changeHeader(event) {
  Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    if (body.text.length == 0)
    {
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
      header.clear();
      firstPageHeader.clear();
      header.insertParagraph("Public - The data is for the public and shareable externally", "Start");
      firstPageHeader.insertParagraph("Public - The data is for the public and shareable externally", "Start");
      header.font.color = "#07641d";
      firstPageHeader.font.color = "#07641d";

      await context.sync();
    }
    else
    {
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
      header.clear();
      firstPageHeader.clear();
      header.insertParagraph("High Confidential - The data must be secret or in some way highly critical", "Start");
      firstPageHeader.insertParagraph("High Confidential - The data must be secret or in some way highly critical", "Start");
      header.font.color = "#f8334d";
      firstPageHeader.font.color = "#f8334d";
      await context.sync();
    }
  });

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

async function insertTable(event) {
  // Implement your custom code here. The following code is a simple Excel example.
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItemOrNullObject("Sample").delete();
      const sheet = context.workbook.worksheets.add("Sample");

      let expensesTable = sheet.tables.add("A1:E1", true);
      expensesTable.name = "SalesTable";
      expensesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];

      expensesTable.rows.add(null, [
        ["Frames", 5000, 7000, 6544, 4377],
        ["Saddles", 400, 323, 276, 651],
        ["Brake levers", 12000, 8766, 8456, 9812],
        ["Chains", 1550, 1088, 692, 853],
        ["Mirrors", 225, 600, 923, 544],
        ["Spokes", 6005, 7634, 4589, 8765],
      ]);

      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();

      sheet.activate();
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    //console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

async function insertImage(event) {
  try {
    await PowerPoint.run(function (context) {
      let shapes = context.presentation.slides.getItemAt(0).shapes;
      const shapeOptions = {
        left: 100,
        top: 100,
        height: 300,
        width: 300
      };
      const hexagon = shapes.addGeometricShape(PowerPoint.GeometricShapeType.hexagon, shapeOptions);
      hexagon.name = "Hexagon";

      shapes = context.presentation.slides.getItemAt(0).shapes;
      const braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair, {
        left: 175,
        top: 450,
        height: 50,
        width: 150
      });
      braces.name = "Braces";
      braces.textFrame.textRange.text = "Shape text";
      braces.textFrame.textRange.font.color = "purple";
      braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;

      return  context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    //console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("changeHeader", changeHeader);
Office.actions.associate("insertTable", insertTable);
Office.actions.associate("insertImage", insertImage);
Office.actions.associate("action", action);
