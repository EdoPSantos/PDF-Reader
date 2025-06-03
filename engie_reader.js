const compactItemRegex = /^(\d{5})(\d{8})([A-ZÇÁÉÍÓÚÀÃÕ\-\s0-9]+?)(\d+)(UN|un|kg|pcs)(\d{4}-\d{2}-\d{2})$/;

function engieReader(lines, STOP_WORDS) {
  const results = [];
  let currentItem = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const match = line.match(compactItemRegex);

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
    }
  }

  if (currentItem) results.push(currentItem);
  return results;
}

module.exports = engieReader;