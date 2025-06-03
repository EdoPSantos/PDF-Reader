const fusedLineRegex = /^(\d{5})(\d{1,3}),\d{2}(un|kg|pcs)(\d{1,3}),\d{2}(\d{1,3}),\d{2}(\d{3})\s*-\s*(.+?)\s*-\s*(\d{2}\/\d{2}\/\d{4})$/i;

function qfReader(lines) {
  const results = [];

  for (const line of lines) {
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
    }
  }

  return results;
}

module.exports = qfReader;
