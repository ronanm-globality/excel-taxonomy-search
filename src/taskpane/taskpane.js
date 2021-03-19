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

// Should be in sync with the default selected value in the html.
let currentSearchType = "office";

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("searchButton").onclick = runSearch;
    document.addEventListener("click", handleGlobalClick, false);
    document.getElementById('search-type-select').addEventListener('change', searchTypeUpdated);
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
  let strategy = findSearchStrategy(currentSearchType);
  var results = await strategy.searchFunction(searchTerm);
  // var countries = await getCountries(searchTerm);

  // Maybe make each of the functions get their own data out first.
  return results;
}

function findSearchStrategy(taxonomyType) {
  const strategies = {
    office: {searchFunction: getOfficeLocations},
    country: {searchFunction: getCountries},
    it_package: {searchFunction: getItPackages},
  }

  if (!(taxonomyType in strategies)) {
    // TODO: Handle this case better.
    return null;
  }

  return strategies[taxonomyType];
}

async function placeResultInTargetCell(context, value) {
  const range = context.workbook.getSelectedRange();
  range.values = [[value]];
  await context.sync();
}

async function handleGlobalClick(event) {
  if (event.target.matches('.searchResult')) {
    await handleSearchResultClick(event);
  } else if (event.target.matches('.options')) {
    toggleOptions();
  }
}

async function handleSearchResultClick(event) {
  try {
    await Excel.run(async context => {
      var outputType = document.querySelector('input[name=outputType]:checked').value;
      var output = (outputType=="uri") ? event.target.getAttribute('data-uri') : event.target.getAttribute('data-name');
      await placeResultInTargetCell(context, output);
    });
  } catch (error) {
    console.error(error);
  }
}

async function getCountries(searchTerm) {
  var url = "https://levant.dev.globality.io/api/v1/country?limit=10&suggestion=";
  var response = await fetch(url + searchTerm);
  var json = await response.json();

  return json.items.map(function (item) {
    return {uri: item.uri, label: item.label}
  });
}

async function getItPackages(searchTerm) {
  var url = "https://levant.test.globality.io/api/v1/it_package?limit=10&suggestion=";
  var response = await fetch(url + searchTerm);
  var json = await response.json();

  return json.items.map(function (item) {
    return {uri: item.uri, label: item.label}
  });
}

async function getOfficeLocations(searchTerm) {
  var url = "https://levant.dev.globality.io/api/v1/place?smaller_than_country=true&limit=10&suggestion=";
  var response = await fetch(url + searchTerm);
  var json = await response.json();
  return json.items.map(function (item) {
    return {uri: item.uri, label: item.qualifiedName}
  });
}

async function displaySearchResults(results) {
  disableSpinner();

  let targetList = document.getElementById("resultsTarget");

  for(var i = 0; i < results.length; i++) {
    var result = results[i];
    var li = document.createElement("li");
    li.appendChild(document.createTextNode(result.label));
    li.setAttribute("class", "searchResult");
    li.setAttribute("data-name", result.label);
    li.setAttribute("data-uri", result.uri);
    targetList.appendChild(li);
  }

  // No Results
  if (results.length === 0) {
    li = document.createElement("li");
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
  // Could change this to use the hidden class like below.
  document.getElementById("loader").style.display = "none";
}

function toggleOptions() {
  document.getElementById("options-content").classList.toggle("hidden");
  document.getElementById("options-caret-down").classList.toggle("hidden");
  document.getElementById("options-caret-up").classList.toggle("hidden");
}

function searchTypeUpdated(update) {
  currentSearchType = update.target.value;
}
