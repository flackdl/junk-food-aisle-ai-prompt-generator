// example spreadsheet
// https://docs.google.com/spreadsheets/d/12UrkZbpyQrOdmO8coUjcb2GsPf0jFSwF3mpePA4w-dg/edit#gid=0

const ExcelJS = window.ExcelJS;
const aq = window.aq;

const resultsButtonsEl = document.getElementById("results-buttons");
const resultsContentEl = document.getElementById("results-content");

function copyToClipboard() {
  const results = getResultsArray().join(' + ');
  navigator.clipboard.writeText(results);
}

function addButton(buttonName, containerEl, className) {
  const newButton = document.createElement("button");
  newButton.classList.add(className);
  newButton.textContent = buttonName;
  containerEl.append(newButton);
  return newButton;
}

function parseSpreadsheet(event) {
  // parse the spreadsheet on file upload and generate column buttons
  const files = event.target.files;
  const reader = new FileReader();
  reader.onload = (e) => {
    
    let dt;
    const workbook = new ExcelJS.Workbook();
    
    // load workbook
    workbook.xlsx
      .load(e.target.result)
      .then(() => {
        return workbook.csv.writeBuffer().then((data) => {
          dt = aq.fromCSV(data.toString());
        });
      })
      .then(() => {
        // display results container
        document.getElementById("results").style = "display: block";

        // add column generator buttons
        const columns = dt.columnNames();
        columns.forEach((col) => {
          // handle column button clicks
          addButton(col, resultsButtonsEl, 'column-button').addEventListener("click", () => {
            
            const existingColumnValues = getExistingValuesForColumn(col);
            
            // generate unique random value
            const itemValue = getRandomItemFromColumn(dt, col, existingColumnValues);
            
            // add item button
            const button = addButton(itemValue, resultsContentEl, 'item-button');            
            
            button.setAttribute('column', col);                        
            
            // handle item replacements            
            button.addEventListener("click", () => {
              button.innerHTML = getRandomItemFromColumn(dt, col, existingColumnValues);
            });
          });
        });
      });
  };
  reader.readAsArrayBuffer(files[0]);
}

function getResultsArray(column) {
  const results = [];  
  Array.from(resultsContentEl.getElementsByTagName('button')).forEach((button) => {
    results.push(button.textContent);
  });
  return results;
}

function getExistingValuesForColumn(column) {
  const results = [];  
  Array.from(resultsContentEl.getElementsByTagName('button')).forEach((button) => {
    if (button.getAttribute('column') === column) {
      results.push(button.textContent);  
    }    
  });
  return results;
}

function resetResults() {
  resultsContentEl.innerHTML = '';
}

function getRandomItem(items) {
  return items[Math.floor(Math.random() * items.length)];
}

function getRandomItemFromColumn(dt, columnName, excludeItems) {
  excludeItems = excludeItems || [];
  return getRandomItem(dt.array(columnName).filter((x) => !excludeItems.includes(x)));
}
