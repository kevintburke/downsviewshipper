<h>K@D Shipper</h>

Processes spreadsheets for K@D SHIP workflows to separate by missing/loaned, mono/serial, and workflow 1/2. Requires ollama and QWEN 2.5 coder.

<em>NOTE: This program is designed for use with items that have descriptions but lack data in the enumeration and chronology fields!</em>

<strong>Process</strong>
<ul>
<li>Ensure script and prompt file are stored in the same directory</li>
<li>Ensure ollama is active and model is installed (type "ollama run qwen2.5-coder" in the command line to activate)</li>
<li>Run the script</li>
<li>Select the source file to process. The file must contain the following columns (headers and order may change):</li>
<ol>
  <li>University of Toronto MMS ID</li>
  <li>University of Ottawa MMS ID</li>
  <li>Title</li>
  <li>Call Number</li>
  <li>Barcode</li>
  <li>Description</li>
  <li>Material Type</li>
  <li>Workflow</li>
  <li>Item Missing</li>
  <li>Item on Loan</li>
  <li>Item ID</li>
</ol>
  <li>Select the appropriate header for each column. All eleven must be filled out.</li>
  <li>Click run</li>
</ul>

The script will confirm that all inputs are present and that the input file is in the correct format before running. The script will follow the process below:
<ol>
  <li>Items on loan will be removed from the dataset if the "on loan" column selected contains any content (is not null)</li>
  <li>Missing items will be removed from the main dataset and added to a spreadsheet for missing items</li>
  <li>Remaining items will have their descriptions parsed using qwen (run locally through ollama). The model will accept the strings of item descriptions and return a Python dictionary with data for Enum A, Enum B, Chron I, and Chron J. Model output will be verified and model will be called again if output fails verification</li>
  <li>Items will be separated by material type (data containing "issue" or "looseleaf" will be classed as serials, all others as monos)</li>
  <li>Items of each material type will be divided by workflow (1 or 2, using the final character of the workflow column selected by the user)</li>
  <li>Data will be cleaned and ordered, with default values added for Library Code, Location Code, and Material Type</li>
  <li>The user will be prompted to save each output (outputs with no content will be skipped)</li>
</ol>

<strong>General Notes:</strong>
<ul>
  <li>The missing and on loan columns will be treated as true/false based on whether they contain any content (including whitespace)</li>
  <li>The material type parsing was developed using Alma's OTB material types</li>
  <li>On a laptop running Intel Core Ultra 5 236V and with 16GB RAM, each call to qwen takes approximately four seconds to complete</li>
</ul>
