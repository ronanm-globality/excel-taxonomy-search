/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

// TODO:
// - Need a choice between value we want to output, like Label or URI.
// - Need choice(?) of different taxonomies? Display values need to be different per each type.
// - Need to update the logo to something Globality.

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("searchButton").onclick = runSearch;
    document.addEventListener("click", handleGlobalClick, false);
  }
});

async function runSearch() {
  try {
    await Excel.run(async context => {

      if (!await isSingleCellSelected(context)) {
        return
      }
      var value = await getValueFromSingleSelectedCell(context);

      // TODO: Maybe bail if there is no search value?

      clearOldSearchResults();
      enableSpinner();

      var results = await search(value);

      await displaySearchResults(results);

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

async function search(searchTerm) {
  // Maybe delegate to the different types of search here?
  var results = await getOfficeLocations(searchTerm);
  // var countries = await getCountries(searchTerm);
  return results.items;
}

async function placeResultInTargetCell(context, value) {
  const range = context.workbook.getSelectedRange();
  range.values = [[value]];
  await context.sync();
}

async function handleGlobalClick(event) {
  if (event.target.matches('.searchResult')) {
    await handleSearchResultClick(event);
  }
}

async function handleSearchResultClick(event) {
  try {
    await Excel.run(async context => {
      var name = event.target.getAttribute('data-name');
      await placeResultInTargetCell(context, name);
    });
  } catch (error) {
    console.error(error);
  }
}

async function getCountries(searchTerm) {
  var url = "https://levant.dev.globality.io/api/v1/country?limit=10&suggestion=";
  var response = await fetch(url + searchTerm);
  return response.json();
}

async function getOfficeLocations(query) {
  var url = "https://levant.dev.globality.io/api/v1/place?smaller_than_country=true&limit=10&suggestion=";
  var response = await fetch(url + query);
  return response.json();
}

async function displaySearchResults(results) {
  disableSpinner();

  let targetList = document.getElementById("resultsTarget");

  for(var i = 0; i < results.length; i++) {
    var result = results[i];
    var li = document.createElement("li");
    li.appendChild(document.createTextNode(result.qualifiedName));
    li.setAttribute("class", "searchResult");
    li.setAttribute("data-name", result.qualifiedName);  // Maybe something better than label here?
    li.setAttribute("data-uri", result.uri);
    targetList.appendChild(li);
  }

  // No Results
  if (results.length === 0) {
    var li = document.createElement("li");
    li.appendChild(document.createTextNode("No Results.."));
    targetList.appendChild(li);
  }
}

function clearOldSearchResults() {
  let targetList = document.getElementById("resultsTarget");
  targetList.innerHTML = "";
}

function enableSpinner() {
  document.getElementById("loader").style.display = "block";
}

function disableSpinner() {
  document.getElementById("loader").style.display = "none";
}
