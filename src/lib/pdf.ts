import jsPDF from "jspdf";

// Paleta institucional (RGB)
export const NAVY: [number, number, number] = [26, 45, 84];
export const NAVY_LIGHT: [number, number, number] = [232, 236, 244];
export const GREY: [number, number, number] = [110, 118, 132];
export const LINE: [number, number, number] = [200, 206, 216];

export const ORGAO = "TRANSALVADOR";
export const ORGAO_SUB = "Superintendência de Trânsito do Salvador";
export const SETOR = "SEGEP — Gerência de Operações Especiais";

/** Cabeçalho institucional. Retorna o Y a partir do qual o conteúdo pode começar. */
export const drawHeader = (doc: jsPDF, titulo: string, subtitulo?: string) => {
  const w = doc.internal.pageSize.getWidth();

  doc.setFillColor(NAVY[0], NAVY[1], NAVY[2]);
  doc.rect(0, 0, w, 26, "F");

  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.text(ORGAO, w / 2, 12, { align: "center" });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.text(ORGAO_SUB, w / 2, 18, { align: "center" });
  doc.text(SETOR, w / 2, 22.5, { align: "center" });

  doc.setFillColor(NAVY_LIGHT[0], NAVY_LIGHT[1], NAVY_LIGHT[2]);
  doc.rect(0, 26, w, 1.6, "F");

  doc.setTextColor(NAVY[0], NAVY[1], NAVY[2]);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12.5);
  doc.text(titulo.toUpperCase(), w / 2, 38, { align: "center" });

  let y = 43;
  if (subtitulo) {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(GREY[0], GREY[1], GREY[2]);
    doc.text(subtitulo, w / 2, y, { align: "center" });
    y += 5;
  }

  doc.setDrawColor(LINE[0], LINE[1], LINE[2]);
  doc.setLineWidth(0.4);
  doc.line(14, y, w - 14, y);
  doc.setTextColor(0, 0, 0);
  return y + 7;
};

/** Bloco de informações em duas colunas com fundo suave. */
export const drawInfoBox = (doc: jsPDF, y: number, items: Array<[string, string]>) => {
  const w = doc.internal.pageSize.getWidth();
  const cols = 2;
  const rows = Math.ceil(items.length / cols);
  const boxH = rows * 6 + 6;

  doc.setFillColor(248, 249, 251);
  doc.setDrawColor(LINE[0], LINE[1], LINE[2]);
  doc.setLineWidth(0.3);
  doc.roundedRect(14, y, w - 28, boxH, 1.5, 1.5, "FD");

  const colW = (w - 28) / cols;
  items.forEach(([label, value], i) => {
    const c = Math.floor(i / rows);
    const r = i % rows;
    const x = 18 + c * colW;
    const ty = y + 8 + r * 6;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.5);
    doc.setTextColor(GREY[0], GREY[1], GREY[2]);
    doc.text(`${label}:`, x, ty);
    const lw = doc.getTextWidth(`${label}: `);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(20, 20, 20);
    doc.text(String(value ?? ""), x + lw, ty);
  });

  return y + boxH + 7;
};

export const tableTheme = {
  headStyles: {
    fillColor: NAVY,
    textColor: [255, 255, 255] as [number, number, number],
    fontStyle: "bold" as const,
    halign: "center" as const,
    fontSize: 8.5,
    cellPadding: 2.2,
  },
  bodyStyles: {
    fontSize: 8.5,
    cellPadding: 2,
    textColor: [30, 30, 30] as [number, number, number],
  },
  alternateRowStyles: { fillColor: [246, 248, 251] as [number, number, number] },
  styles: { lineColor: LINE, lineWidth: 0.2, font: "helvetica" },
  footStyles: {
    fillColor: NAVY_LIGHT,
    textColor: NAVY,
    fontStyle: "bold" as const,
    fontSize: 8.5,
  },
};

/** Assinaturas lado a lado. */
export const drawSignatures = (
  doc: jsPDF,
  y: number,
  blocos: Array<{ titulo: string; sub?: string }>
) => {
  const w = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  if (y > pageH - 42) {
    doc.addPage();
    y = 45;
  }
  const usable = w - 28;
  const each = usable / blocos.length;
  blocos.forEach((b, i) => {
    const cx = 14 + each * i + each / 2;
    doc.setDrawColor(80, 80, 80);
    doc.setLineWidth(0.4);
    doc.line(cx - each / 2 + 12, y, cx + each / 2 - 12, y);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.5);
    doc.setTextColor(20, 20, 20);
    doc.text(b.titulo, cx, y + 5, { align: "center" });
    if (b.sub) {
      doc.setFont("helvetica", "normal");
      doc.setTextColor(GREY[0], GREY[1], GREY[2]);
      doc.setFontSize(7.5);
      doc.text(b.sub, cx, y + 9.5, { align: "center" });
    }
  });
  return y + 16;
};

/** Rodapé com paginação — chamar por último. */
export const drawFooters = (doc: jsPDF, contato?: string) => {
  const w = doc.internal.pageSize.getWidth();
  const h = doc.internal.pageSize.getHeight();
  const total = doc.getNumberOfPages();
  const emissao = new Date().toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" });
  for (let p = 1; p <= total; p++) {
    doc.setPage(p);
    doc.setDrawColor(LINE[0], LINE[1], LINE[2]);
    doc.setLineWidth(0.3);
    doc.line(14, h - 14, w - 14, h - 14);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(GREY[0], GREY[1], GREY[2]);
    doc.text(`GEOPS — emitido em ${emissao}`, 14, h - 9.5);
    if (contato) doc.text(contato, w / 2, h - 9.5, { align: "center" });
    doc.text(`Página ${p} de ${total}`, w - 14, h - 9.5, { align: "right" });
  }
};
