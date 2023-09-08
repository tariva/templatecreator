const { create } = require("domain");
const fs = require("fs");
const XLSX = require("xlsx");
const iconv = require("iconv-lite");

const directoryPath = "./testfiles/";
const outputPath = "./outputfiles/";
// Delete and recreate the outputfiles directory

try {
  fs.rmdirSync("outputfiles", { recursive: true });
  console.log("Successfully deleted the outputfiles directory.");
} catch (err) {
  console.error("Error while deleting the outputfiles directory.", err);
}

try {
  fs.mkdirSync("outputfiles");
  console.log("Successfully created the outputfiles directory.");
} catch (err) {
  console.error("Error while creating the outputfiles directory.", err);
  return; // Exit if there's an error
}

// Read all files in the directory
fs.readdir(directoryPath, (err, files) => {
  if (err) {
    return console.error("Unable to scan directory: " + err);
  }
  const wb = XLSX.utils.book_new();
  let data = [];
  let keys = [];
  let values = [];
  let processedFiles = [];
  files.forEach((file) => {
    // Ignore files with extensions
    if (!file.includes(".")) {
      console.log("Reading file: " + file);
      const contentBuffer = fs.readFileSync(directoryPath + file);
      const content = iconv.decode(contentBuffer, "win1252");
      const lines = content.split("\n");
      processedFiles.push(file);
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
    if (data.length > 0) {
      // Create worksheet
      const ws = XLSX.utils.aoa_to_sheet(data);

      // Write data to Excel

      XLSX.utils.book_append_sheet(wb, ws, file);
      // write to output folder
    }
  });
  XLSX.writeFile(wb, "outputfiles/output.xlsx");
  console.log("Excel file created as output.xlsx");
  // After processing all files, write the names of processed files to a text file
  fs.writeFileSync("processed_files.txt", processedFiles.join("\n"));
});
