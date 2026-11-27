import type { Feature, SheetData } from "write-excel-file/browser";

export type ExportRow = {
  name: string;
  dwt_k: number | null;
  built: number | null;
  open_location: string | null;
  open_country?: string | null;
  zone: string;
  open_date: string | null;
  received: string;
  from: string;
  subject: string;
  webLink: string;
};

const HEADERS = ["Vessel", "DWT (K)", "Built", "Open", "Country", "Zone", "Open date", "Received", "From", "Email"];
const WIDTHS = [26, 9, 7, 22, 9, 10, 14, 17, 30, 60];

// Excel's filter/sort dropdowns on the header row (not built into write-excel-file).
function autoFilter(lastRow: number): Feature<File | Blob | ArrayBuffer> {
  const ref = `A1:${String.fromCharCode(64 + HEADERS.length)}${lastRow}`;
  return {
    files: {
      transform: {
        "xl/worksheets/sheet{id}.xml": {
          transform: (xml) => (xml.includes("<autoFilter") ? xml : xml.replace("</sheetData>", `</sheetData><autoFilter ref="${ref}"/>`)),
        },
      },
    },
  };
}

/** Build and download a formatted .xlsx (frozen header, filters, numbers as numbers, dates as dates). */
export async function downloadXlsx(rows: ExportRow[], fileName: string) {
  const { default: writeXlsxFile } = await import("write-excel-file/browser"); // loaded only when needed
  const header = HEADERS.map((h) => ({ value: h, fontWeight: "bold" as const }));
  const data: SheetData = [
    header,
    ...rows.map((r) => [
      { value: r.name, type: String },
      r.dwt_k != null ? { value: r.dwt_k, type: Number } : null,
      r.built != null ? { value: r.built, type: Number } : null,
      r.open_location ? { value: r.open_location, type: String } : null,
      r.open_country ? { value: r.open_country, type: String } : null,
      r.zone ? { value: r.zone, type: String } : null,
      r.open_date ? { value: r.open_date, type: String } : null,
      { value: new Date(r.received), type: Date, format: "dd mmm yyyy hh:mm" },
      { value: r.from, type: String },
      { value: r.subject || "(no subject)", type: String },
    ]),
  ];
  await writeXlsxFile(
    data,
    { sheet: "Positions", stickyRowsCount: 1, columns: WIDTHS.map((width) => ({ width })) },
    { features: [autoFilter(rows.length + 1)] },
  ).toFile(fileName);
}
