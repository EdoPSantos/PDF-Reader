const glnBlockPattern = [
  /^\d{1,3},\d{3}$/,
  /^[A-Z]{2}$/,
  /^\d{3,5}$/,
  /^.+$/,
  /^.+$/,
  /^.+$/,
];

function gnlReader(lines) {
  const results = [];

  for (let i = 0; i < lines.length; i++) {
    const block = lines.slice(i, i + 6);
    const isGnl = block.length === 6 && block.every((line, idx) => glnBlockPattern[idx].test(line));

    if (isGnl) {
      const [quantityRaw, prefix, destination, designationRaw, specsRaw, partRefRaw] = block;

      const quantityMatch = quantityRaw.match(/^(\d{1,3}),(\d{3})$/);
      if (!quantityMatch) continue;

      const quantity = parseInt(quantityMatch[1], 10);
      const itemNumber = parseInt(quantityMatch[2], 10);

      let designation = designationRaw.trim();
      let specifications = specsRaw.trim();

      if (!/^[A-Z]+\d+$/.test(designation)) {
        const parts = designation.split(/\s+|-/);
        if (parts.length > 1 && /^[A-Z]+$/.test(parts[0])) {
          designation = parts[0].trim();
          specifications = parts.slice(1).join(" ").trim();
        }
      }

      let parts = partRefRaw.trim();
      let reference = null;

      const refMatch = partRefRaw.match(/^(.+?)\s*\(?(\d+)\)?$/);
      if (refMatch) {
        parts = refMatch[1].trim();
        reference = `(${refMatch[3]})`;
      }

      results.push({
        "Item Number": itemNumber,
        Destiny: destination,
        Quantity: quantity,
        Designation: designation,
        Specifications: specifications,
        Parts: parts,
        Reference: reference
      });

      i += 5;
    }
  }

  return results;
}

module.exports = gnlReader;
