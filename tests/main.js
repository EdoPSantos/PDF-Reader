const fs = require("fs");
const pdfParse = require("pdf-parse");
const readline = require("readline");

const engieReader = require("./engie_reader");
const mdReader = require("./md_group_reader");
const reromReader = require("./gln_reader");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

const options = {
  option1: "./pdfs/opcao1.pdf",   // ENGIE
  option2: "./pdfs/opcao2.PDF",   // MD GROUP
  option3: "./pdfs/opcao3.PDF",   // GLN
};

console.log("Escolha o PDF para processar:");
console.log("1. ENGIE");
console.log("2. MD GROUP");
console.log("3. GLN");
console.log("4. TODOS OS ACIMA\n");

rl.question("Digite o número da opção: ", async (answer) => {
  const STOP_WORDS = [
    "Item", "Material", "Descrição", "Solicitação", "ENGIE",
    "Processado por computador", "Emitido por", "Pág.", "Data de emissão",
    "Condições Gerais", "Email", "N°", "A presente Solicitação de Proposta",
    "PG.33.001.PRT", "disponíveis em", "Com a resposta a esta solicitação",
    "O nosso número fiscal de identificação",
    "as condições de entrega indicadas nesta Solicitação",
    "não carece de assinatura", "Name:", "Emitido por:"
  ];

  async function processPDF(path, readerFn, label, stopWords = null) {
    if (!fs.existsSync(path)) {
      console.warn(`Ficheiro não encontrado: ${path}`);
      return [];
    }

    const buffer = fs.readFileSync(path);
    const data = await pdfParse(buffer);

    const lines = data.text
      .split("\n")
      .map(l => l.trim())
      .filter(l => l !== "");

    const result = stopWords ? readerFn(lines, stopWords) : readerFn(lines);

    console.log(`\nResultados de ${label}:\n`);
    console.table(result);

    if (label === "GLN") {
      console.log("Ainda não está 100% funcional\n");
    }

    return result;
  }

  if (answer === "4") {
    await processPDF(options.option1, engieReader, "ENGIE", STOP_WORDS);
    await processPDF(options.option2, mdReader, "MD GROUP");
    await processPDF(options.option3, reromReader, "GLN");
  } else {
    const selectedKey = `option${answer}`;
    const selectedFile = options[selectedKey];

    if (!selectedFile || !fs.existsSync(selectedFile)) {
      console.error("Opção inválida ou ficheiro não encontrado.");
      rl.close();
      return;
    }

    const buffer = fs.readFileSync(selectedFile);
    const data = await pdfParse(buffer);

    const lines = data.text
      .split("\n")
      .map(l => l.trim())
      .filter(l => l !== "");

    let results = [];

    if (answer === "1") {
      results = engieReader(lines, STOP_WORDS);
    } else if (answer === "2") {
      results = mdReader(lines);
    } else if (answer === "3") {
      results = reromReader(lines);
    }

    console.log("\nItens extraídos:\n");
    console.table(results);

    if (answer === "3" || answer === "4") {
      console.log("GLN ainda não está 100% funcional\n");
    }
  }

  rl.close();
});
