const fs = require("fs");
const pdfParse = require("pdf-parse");
const readline = require("readline");

const engieReader = require("./engieReader");
const qfReader = require("./qfReader");
const reromReader = require("./reromReader");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

const options = {
  option1: "../pdf_reader/opcao1.PDF",   // ENGIE
  option2: "../pdf_reader/opcao2.pdf",   // MD GROUP
  option3: "../pdf_reader/opcao3.Pdf",   // GLN
};

console.log("Escolha o PDF para processar:");
console.log("1. ENGIE");
console.log("2. MD GROUP");
console.log("3. GLN\n");

rl.question("Digite o número da opção: ", (answer) => {
  const selectedKey = `option${answer}`;
  const selectedFile = options[selectedKey];

  if (!selectedFile || !fs.existsSync(selectedFile)) {
    console.error("⚠️ Opção inválida ou ficheiro não encontrado.");
    rl.close();
    return;
  }

  const buffer = fs.readFileSync(selectedFile);

  pdfParse(buffer).then(function (data) {
    const lines = data.text
      .split("\n")
      .map(l => l.trim())
      .filter(l => l !== "");

    const STOP_WORDS = [
      "Item", "Material", "Descrição", "Solicitação", "ENGIE",
      "Processado por computador", "Emitido por", "Pág.", "Data de emissão",
      "Condições Gerais", "Email", "N°", "A presente Solicitação de Proposta",
      "PG.33.001.PRT", "disponíveis em", "Com a resposta a esta solicitação",
      "O nosso número fiscal de identificação",
      "as condições de entrega indicadas nesta Solicitação",
      "não carece de assinatura", "Name:", "Emitido por:"
    ];

    const resultsEngie = engieReader(lines, STOP_WORDS);
    const resultsQF = qfReader(lines);
    const resultsRerom = reromReader(lines);

    const allResults = [
      ...resultsEngie,
      ...resultsQF,
      ...resultsRerom
    ];

    console.log("\nItens extraídos:\n");
    console.table(allResults);

     if (answer === "3") {
      console.log("Ainda não está 100% funcional\n");
    }
    
    rl.close();
  });
});