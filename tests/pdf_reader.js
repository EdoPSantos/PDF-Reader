const fs = require("fs");
const path = require("path");
const pdfParse = require("pdf-parse");
const readline = require("readline");
const PDFParser = require("pdf2json");

const folderPath = "../pdfs";

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Lista de PDFs
function listPDFFiles() {
  return fs.readdirSync(folderPath).filter(f => f.toLowerCase().endsWith(".pdf"));
}

// Pergunta qual PDF o utilizador quer abrir
function askUserToChoosePDF(files) {
  console.log("Escolha um dos seguintes ficheiros PDF:");
  files.forEach((f, i) => {
    console.log(`${i + 1}. ${f}`);
  });

  return new Promise((resolve) => {
    rl.question("\nDigite o número da opção: ", (answer) => {
      const index = parseInt(answer) - 1;
      if (index >= 0 && index < files.length) {
        resolve(files[index]);
      } else {
        console.log("Opção inválida.");
        rl.close();
      }
    });
  });
}

// Pergunta qual leitor o utilizador quer usar
function askUserReaderType() {
  console.log("\nQual leitor deseja usar?");
  console.log("1. Leitor simples (pdf-parse)");
  console.log("2. Leitor detalhado (pdf2json)");

  return new Promise((resolve) => {
    rl.question("Digite o número da opção: ", (answer) => {
      const option = parseInt(answer.trim());
      resolve(option);
    });
  });
}

// Leitor simples com pdf-parse
async function readWithPdfParse(filePath) {
  const buffer = fs.readFileSync(filePath);
  const data = await pdfParse(buffer);
  console.log("\n--- Conteúdo do PDF (pdf-parse) ---\n");
  console.log(data.text);
  console.log("\n--- Fim do conteúdo ---\n");
}

// Leitor detalhado com pdf2json
function readWithPdf2Json(filePath) {
  return new Promise((resolve) => {
    const pdfParser = new PDFParser();

    pdfParser.on("pdfParser_dataReady", pdfData => {
  if (!pdfData.formImage || !pdfData.formImage.Pages) {
    console.error("Erro: O ficheiro não contém estrutura reconhecida pelo pdf2json.");
    resolve();
    return;
  }

  const pages = pdfData.formImage.Pages;

  console.log("\n--- Conteúdo detalhado do PDF (pdf2json) ---\n");

  pages.forEach((page, pageIndex) => {
    console.log(`Página ${pageIndex + 1}:\n`);

    page.Texts.forEach(text => {
      const content = decodeURIComponent(text.R.map(r => r.T).join(""));
      const x = text.x;
      const y = text.y;
      console.log(`Texto: "${content}" (x: ${x}, y: ${y})`);
    });

    console.log("\n-------------------------\n");
  });

  resolve();
});

    pdfParser.loadPDF(filePath);
  });
}

async function main() {
  while (true) {
    const pdfFiles = listPDFFiles();

    if (pdfFiles.length === 0) {
      console.log("Nenhum ficheiro PDF encontrado na pasta.");
      break;
    }

    const selectedFile = await askUserToChoosePDF(pdfFiles);
    if (!selectedFile) break;

    const filePath = path.join(folderPath, selectedFile);
    const readerChoice = await askUserReaderType();

    if (readerChoice === 1) {
      await readWithPdfParse(filePath);
    } else if (readerChoice === 2) {
      await readWithPdf2Json(filePath);
    } else {
      console.log("Opção de leitor inválida.");
    }

    let continuar;
    while (true) {
    continuar = await new Promise((resolve) => {
        rl.question("Deseja ver outro ficheiro? (s/n): ", (answer) => {
        resolve(answer.trim().toLowerCase());
        });
    });

    if (continuar === "s" || continuar === "sim" || continuar === "n" || continuar === "não") {
        break;
    } else {
        console.log("Opção inválida. Por favor responda com 's' ou 'n'.");
    }
    }

    if (continuar === "n" || continuar === "não") break;
  }

  rl.close();
}

main();
