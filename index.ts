import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

import * as fs from "fs";
import * as path from "path";

const content = fs.readFileSync(
  path.resolve(__dirname, "input.docx"),
  "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
});

// Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
doc.render({
  nome: "Henrique Souza",
  idade: "27",
  e_legal: "Sim",
  possuiCachorro: true,
  nome_cachorro: "Paçoca",
  tabela_divertida: [
    {
      nome_item_divertido: "Primeiro Item Divertido",
      preco_item_divertido: 10.5,
    },
    {
      nome_item_divertido: "Segundo Item Divertido",
      preco_item_divertido: 12.0,
    },
    {
      nome_item_divertido: "Terceiro Item Divertido",
      preco_item_divertido: 15.212,
    },
  ],
  tecnologias: [
    { tech: "Javascript" },
    { tech: "C#" },
    { tech: "Algo bem maluco e doido" },
  ],
});

const buf = doc.getZip().generate({
  type: "nodebuffer",
  // compression: DEFLATE adds a compression step.
  // For a 50MB output document, expect 500ms additional CPU time
  compression: "DEFLATE",
});

// buf is a nodejs Buffer, you can either write it to a file or res.send it with express for example.
fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);
