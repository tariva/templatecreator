const { create } = require("domain");
const fs = require("fs");
const XLSX = require("xlsx");
const iconv = require("iconv-lite");

const directoryPath = "./testfiles/"; // Current directory. Change this if needed.
// delete and create output folder
fs.rmdirSync("outputfiles", { recursive: true });
fs.mkdirSync("outputfiles");

// Read all files in the directory
fs.readdir(directoryPath, (err, files) => {
  if (err) {
    return console.error("Unable to scan directory: " + err);
  }

  let data = [];
  let keys = [];
  let values = [];
  files.forEach((file) => {
    // Ignore files with extensions
    if (!file.includes(".")) {
      console.log("Reading file: " + file);
      const contentBuffer = fs.readFileSync(directoryPath + file);
      const content = iconv.decode(contentBuffer, "win1252");
      const lines = content.split("\n");

      for (let line of lines) {
        /*
        In this regex:
        (...) captures any three characters.
        =\s* matches an equals sign followed by any number of whitespace characters.
        (.*?) captures any character (non-greedy) until the next pattern.
        \s* matches any number of whitespace characters.
        \[(.*?)\] captures the content between square brackets.
         */
        const match = line.match(/(...)=\s*(.*?)\s*\[(.*?)\];/);

        if (match) {
          console.log(match[1], match[2].trim(), match[3]);
          if (match) {
            keys.push(match[3]);
            values.push(match[2].trim());
          }
        }
      }
      data = [keys, values];
    }
  });

  // Write data to Excel
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  // write to output folder
  XLSX.writeFile(wb, "outputfiles/output.xlsx");
  console.log("Excel file created as output.xlsx");
});
