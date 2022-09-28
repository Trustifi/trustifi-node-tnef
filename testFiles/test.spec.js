
const fs = require('fs');
const {parseBuffer} = require('../src/parse');

const winmaildat = fs.readFileSync('attachments.dat');
let tnefObj = parseBuffer(winmaildat);

console.log(tnefObj);
