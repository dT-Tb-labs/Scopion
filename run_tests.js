// Node harness for Test.gs. Loads the document-free parts of the add-on and
// runs the same assertions Apps Script would run via runTests().
const fs = require('fs');
const vm = require('vm');

const sandbox = { Logger: { log: (m) => console.log(m) }, console, Date, Math, JSON, String, Number, Array, isNaN, parseInt };
vm.createContext(sandbox);

for (const file of ['Formula.gs', 'Trace.gs', 'Test.gs']) {
  vm.runInContext(fs.readFileSync(file, 'utf8'), sandbox, { filename: file });
}
try {
  console.log(sandbox.runTests());
} catch (e) {
  console.error(e.message);
  process.exit(1);
}
