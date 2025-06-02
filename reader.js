const fs = require("fs");
const pdfParse = require("pdf-parse");

// Caminho para o PDF atual
const pathPDF = "../pdf_reader/QF 25_0001037.pdf";
const buffer = fs.readFileSync(pathPDF);

pdfParse(buffer).then(function (data) {
  const text = data.text;
  const lines = text
    .split("\n")
    .map((l) => l.trim())
    .filter((l) => l !== "");

    //console.log("\n--- Conteúdo do PDF (linhas) ---\n");
    //console.log(lines.join("\n"));
    //console.log("\n--- Fim do conteúdo ---\n");

  const results = [];

  const compactItemRegex = /^(\d{5})(\d{8})([A-ZÇÁÉÍÓÚÀÃÕ\-\s0-9]+?)(\d+)(UN|un|kg|pcs)(\d{4}-\d{2}-\d{2})$/;
  const fusedLineRegex = /^(\d{5})(\d{1,3}),\d{2}(un|kg|pcs)(\d{1,3}),\d{2}(\d{1,3}),\d{2}(\d{3})\s*-\s*(.+?)\s*-\s*(\d{2}\/\d{2}\/\d{4})$/i;



  const STOP_WORDS = [
    "Item", "Material", "Descrição", "Solicitação", "ENGIE",
    "Processado por computador", "Emitido por", "Pág.", "Data de emissão",
    "Condições Gerais", "Email", "N°", "A presente Solicitação de Proposta",
    "PG.33.001.PRT", "disponíveis em", "Com a resposta a esta solicitação",
    "O nosso número fiscal de identificação",
    "as condições de entrega indicadas nesta Solicitação",
    "não carece de assinatura", "Name:", "Emitido por:"
  ];

  let currentItem = null;
  let prevLine = "";

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // 1. Layout ENGIE
    let match = line.match(compactItemRegex);
    if (match) {
      if (currentItem) results.push(currentItem);

      const [
        _, itemNumber, materialCode, rawDescription,
        quantity, um, deliveryDate
      ] = match;

      currentItem = {
        "Item Number": itemNumber,
        "Material Code": materialCode,
        Quantity: quantity,
        UM: um.toUpperCase(),
        "Delivery Date": deliveryDate,
        Description: rawDescription.trim(),
        Brand: null,
        Reference: null,
        Supplier: null,
      };

      let j = i + 1;
      while (j < lines.length) {
        const nextLine = lines[j];

        if (
          compactItemRegex.test(nextLine) ||
          STOP_WORDS.some((word) => nextLine.toLowerCase().includes(word.toLowerCase()))
        ) break;

        if (/^Marca:/i.test(nextLine)) {
          currentItem["Brand"] = nextLine.replace(/^Marca:\s*/i, "").trim();
        } else if (/^REF:/i.test(nextLine)) {
          currentItem["Reference"] = nextLine.replace(/^REF:\s*/i, "").trim();
        } else if (/^Fornecedor:/i.test(nextLine)) {
          currentItem["Supplier"] = nextLine.replace(/^Fornecedor:\s*/i, "").trim();
        } else {
          currentItem["Description"] += " " + nextLine;
        }

        j++;
      }

      i = j - 1;
      prevLine = "";
      continue;
    }

    // 2. Layout QF colado
    const fusedMatch = line.match(fusedLineRegex);
    if (fusedMatch) {
    const [
        _, itemNumber, quantityStr, unit, priceStr, amountStr,
        codeExtra, description, deliveryDate
    ] = fusedMatch;

    results.push({
        "Item Number": itemNumber,
        Quantity: parseFloat(quantityStr.replace(",", ".")),
        UM: unit.toUpperCase(),
        "Delivery Date": deliveryDate,
        Description: `MD5863 ${codeExtra} - ${description.trim()}`,
        Price: parseFloat(priceStr.replace(",", ".")),
        Amount: parseFloat(amountStr.replace(",", ".")),
    });

    prevLine = "";
    continue;
    }

    if (
      !line.match(/^\d{5}/) &&
      !STOP_WORDS.some((w) => line.toLowerCase().includes(w.toLowerCase()))
    ) {
      prevLine = line;
    }
  }

  if (currentItem) results.push(currentItem);


  console.log("\nItens extraídos:\n");
  console.table(results);
});
