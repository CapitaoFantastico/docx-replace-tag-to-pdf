"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const pizzip_1 = __importDefault(require("pizzip"));
const docxtemplater_1 = __importDefault(require("docxtemplater"));
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
const content = fs.readFileSync(path.resolve(__dirname, "input.docx"), "binary");
const zip = new pizzip_1.default(content);
const doc = new docxtemplater_1.default(zip, {
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
