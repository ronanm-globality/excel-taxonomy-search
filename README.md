# Globality Taxonomy Search

This project is an add-in for Excel to add a component to be able to perform a search on the internal Globality Taxonomies directly from Excel.

![App Screenshot](https://i.imgur.com/70gaXGY.png)

## How to use it?

1. Populate a cell with the contents of your search.
2. With the cell still selected, hit Search in the taskpane.
3. Results will get populated. Click on a search result to populate the value back into the currently selected cell.
4. Use the options dropdown to toggle between label/uri values, along with picking the type of taxonomy search (default is Office Locations)

## How to run locally?
Install dependencies:
```
npm install
```

There is a built in command to run a local server:
```
npm run dev-server
```
Then there is a command to install the `manifest.xml` file to Excel:
```
npm start
```

This should pop open Excel and you should see the new icon in the top right of the Home option bar.

## Debugging

This template supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)
