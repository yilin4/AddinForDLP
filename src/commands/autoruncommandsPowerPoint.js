/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

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

// The add-in command functions need to be available in global scope

Office.actions.associate("insertImage", insertImage);
