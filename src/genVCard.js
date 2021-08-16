const vCardsJS = require("vcards-js");
const readXlsxFile = require("read-excel-file/node");
fs = require("fs");

let output = "";

readXlsxFile("./src/Contacts.xlsx").then((rows) => {
  rows.forEach((el) => {
    const vCard = vCardsJS();
    const numPhone = String(el[2]);
    vCard.firstName = numPhone;
    vCard.cellPhone = numPhone;
    output += vCard.getFormattedString();
  });
  fs.writeFile("vcard.vcf", output, (err, data) => {
    if (err) console.log(err);
  });
});
