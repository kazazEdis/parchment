import { describe, it, expect } from "vitest";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell } from "docx";
import { readDocx, writeDocx } from "../src/opc";
import { fromPackage, fillTemplate } from "../src/doc";

async function bigDoc(nParas: number): Promise<Uint8Array> {
  const children: (Paragraph | Table)[] = [];
  for (let i = 0; i < nParas; i++) {
    children.push(new Paragraph({ children: [new TextRun(`Paragraph ${i}: offer {n} for {c} totalling {t} EUR.`)] }));
  }
  for (let t = 0; t < 10; t++) {
    children.push(new Table({ rows: [new TableRow({ children: [
      new TableCell({ children: [new Paragraph("A")] }),
      new TableCell({ children: [new Paragraph("B")] }),
    ] })] }));
  }
  return new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children }] })));
}

describe("bench: full headless read → parse → fill → write cycle", () => {
  it("processes a large document fast (and end-to-end correct)", async () => {
    const nParas = 400;
    const bytes = await bigDoc(nParas);

    const t0 = performance.now();
    const pkg = await readDocx(bytes);
    const t1 = performance.now();
    const doc = fromPackage(pkg); // parse document + styles + numbering
    const t2 = performance.now();
    const { count } = fillTemplate(doc, { n: "2026-001", c: "ACME", t: "1.234,56" });
    const t3 = performance.now();
    await writeDocx(doc.pkg);
    const t4 = performance.now();

    // eslint-disable-next-line no-console
    console.log(
      `[docx-bench] ${nParas} paras + 10 tables | read ${(t1 - t0).toFixed(1)}ms | parse ${(t2 - t1).toFixed(1)}ms | ` +
        `fill(${count}) ${(t3 - t2).toFixed(1)}ms | write ${(t4 - t3).toFixed(1)}ms | total ${(t4 - t0).toFixed(1)}ms`,
    );

    expect(count).toBe(nParas * 3); // 3 tokens per paragraph
    expect(t4 - t0).toBeLessThan(5000); // generous regression bound
  });
});

