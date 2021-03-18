/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("place").onclick = pickResult;
  }
});

async function run() {
  try {
    await Excel.run(async context => {

      if (!await isSingleCellSelected(context)) {
        return
      }
      var value = await getValueFromSingleSelectedCell(context);

      var searchResults = await search(value);
      console.log(searchResults);

    });
  } catch (error) {
    console.error(error);
  }
}

async function pickResult() {
  try {
    await Excel.run(async context => {
      // Does it just go back to the selected cell? Or the original one?
      var pickedResult = "ronan";

      if (!await isSingleCellSelected(context)) {
        return
      }

      await placeResultInTargetCell(context, pickedResult);

    });
  } catch (error) {
    console.error(error);
  }
}

async function isSingleCellSelected(context) {
  const range = context.workbook.getSelectedRange();
  range.load("cellCount");
  await context.sync();
  return range.cellCount === 1;
}

async function getValueFromSingleSelectedCell(context) {
  const range = context.workbook.getSelectedRange();
  range.load("values");
  await context.sync();
  // Values are returned as a 2D array. We've already checked it's a single value.
  return range.values[0][0];
}

async function search(value) {
  console.log("searching for:", value);
  return ["Ireland"]
}

async function placeResultInTargetCell(context, value) {
  const range = context.workbook.getSelectedRange();
  range.values = [[value]];
  await context.sync();
}
