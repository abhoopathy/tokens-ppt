const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");

const src = "theme_extracted";
const out = "mytheme.thmx";

if (!fs.existsSync(src)) {
  console.error(`ERROR: folder '${src}' not found.`);
  console.log("Files here:", fs.readdirSync("."));
  process.exit(1);
}

const zip = new AdmZip();
const filesFound = [];

function walk(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const abs = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      walk(abs);
    } else {
      const arcPath = path.relative(src, abs).replace(/\\/g, "/");
      zip.addFile(arcPath, fs.readFileSync(abs));
      filesFound.push(arcPath);
    }
  }
}

walk(src);
zip.writeZip(out);

console.log(`Written ${filesFound.length} files to ${out}:`);
filesFound.forEach(f => console.log(" ", f));