import axios from 'axios';
import Docxtemplater from 'docxtemplater';
import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import { doSubjectList } from './url.js';

const timestamp = Date.parse(new Date());

const request = {
  "courseid": "MDAwMDAwMDAwMLOGsZmHuc1shKVyoQ",
  "reqtimestamp": 1635584013370,
  "testpaperid": "MDAwMDAwMDAwMLOGvZaHz7tthNtyoQ"
}
axios.defaults.responseType='json';

const header = {
 Accept: 'application/json, text/plain, */*',
 ContentType: 'application/json',
Host: 'openapiv53.ketangpai.com',
Origin: 'https://www.ketangpai.com',
Referer: 'https://www.ketangpai.com/',
token: 'ebb2a4db5c247fa11938b55e08f0074cd7ac76a1279b54140039c23d43edb557',
}


function main () {
    axios.post(doSubjectList, request,
      {
        headers:header,
        proxy:false
      }
    ).then(response => {
    dispose(response.data.data)
  }).catch(err => {
    console.log(err)
  })
}
main()

function dispose (data) {
// Load the docx file as binary content
let content = fs.readFileSync(
    path.resolve( "tag-example.docx"),
    "binary"
);

let zip = new PizZip(content)

function parser(tag) {
  return {
      get(scope, context) {
          if (tag === "$index") {
              const indexes = context.scopePathItem;
              return indexes[indexes.length - 1] + 1;
          }
          return scope[tag];
      },
  };
}
const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    parser
});
doc.render({
    title: data.testpaper.title,
    lists: data.lists
});

let buf = doc.getZip().generate({ type: "nodebuffer" });
// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
fs.writeFileSync(path.resolve( "output.docx"), buf);
}