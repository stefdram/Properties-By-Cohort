// JSON parsing file

const fs = require('fs');

const pbcJsonParser = () => {
  const rawData = fs.readFileSync('propertiesByCohort.json');
  const propertiesByCohort = JSON.parse(rawData);
  return propertiesByCohort;
}

module.exports = pbcJsonParser;