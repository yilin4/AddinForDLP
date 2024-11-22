/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

async function changeHeader(event) {
  Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    if (body.text.length == 0)
    {
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      header.clear();
      header.insertParagraph("Public - The data is for the public and shareable externally", "Start");
      const font = header.font;
      font.color = "#07641d";

      await context.sync();
    }
    else
    {
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      header.clear();
      header.insertParagraph("High Confidential - The data must be secret or in some way highly critical", "Start");
      const font = header.font;
      font.color = "#f8334d";
      await context.sync();
    }
  });

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope

Office.actions.associate("changeHeader", changeHeader);