import { Document, Packer, Paragraph, HeadingLevel, Table, TableCell, TableRow, WidthType, Alignment, AlignmentType, ShadingType } from "docx";
import { saveAs } from "file-saver";
import * as fs from "fs";

let table = new Table({
    columnWidths: [7208, 2000],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [new Paragraph({text:"Warning !!! กรุณาตรวจสอบในหัวข้อนี้ให้ครบถ้วนก่อนส่งงานมายัง ACTM", style: "textRedColor"})],
                    shading: {
                        fill: "fef010",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [new Paragraph({ text: "*Yes/No", style: "textRedColor", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fef010",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [new Paragraph({text:"    1. ทำการเปลี่ยน Config, Path, User, Password จากระบบ Test เป็นระบบ Production แล้วหรือไม่", style: "textNomal"})],
                    shading: {
                        fill: "fffc86",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [new Paragraph({ text: "Yes", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fffc86",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [new Paragraph({text:"    2. ทำการลบ Comment, Remark ออกแล้วหรือไม่", style: "textNomal"})],
                    shading: {
                        fill: "fffc86",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [new Paragraph({ text: "Yes", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fffc86",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [new Paragraph({text:"    3. File เป็น Version ล่าสุดหรือไม่", style: "textNomal"})],
                    shading: {
                        fill: "fffc86",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [new Paragraph({ text: "Yes", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fffc86",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [],
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.PERCENTAGE,
                    },
                    children: [],
                }),
            ],
        }),
    ],
});

function saveDocumentToFile(doc, fileName) {
  // Create new instance of Packer for the docx module

  // Create a mime type that will associate the new file with Microsoft Word
  const mimeType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  // Create a Blob containing the Document instance and the mimeType
  Packer.toBlob(doc).then((blob) => {
    const docblob = blob.slice(0, blob.size, mimeType);
    // Save the file using saveAs from the file-saver package
    saveAs(docblob, fileName);
  });
}

function generateWordDocument(event) {
  event.preventDefault();
  // Create a new instance of Document for the docx module

  let doc = new Document({
    styles: {
      paragraphStyles: [
        {
          id: "myCustomStyle",
          name: "My Custom Style",
          basedOn: "Normal",
          run: {
            color: "FF0000",
            italics: true,
            bold: true,
            size: 26,
            font: "Calibri"
          },
          paragraph: {
            spacing: { line: 276, before: 150, after: 150 }
          }
        },
        {
            id: "textHeader",
            name: "Text Header",
            basedOn: "Normal",
            run: {
              color: "030303",
              bold: true,
              size: 26,
              font: "Sarabun"
            }
        },
        {
            id: "textNomal",
            name: "Text Nomal",
            basedOn: "Normal",
            run: {
              color: "030303",
              size: 22,
              font: "Sarabun"
            }
        },
        {
            id: "textRedColor",
            name: "Text Red Color",
            basedOn: "Normal",
            run: {
              color: "FF0000",
              size: 22,
              font: "Sarabun"
            }
        }
      ]
    },
    sections: [
      {
        children: [
          new Paragraph({ text: "Warning (for Solutions)",
          style: "textHeader",
          alignment: AlignmentType.CENTER
        }),
          table,
        ]
      }
    ]
  });
  // Call saveDocumentToFile with the document instance and a filename
  saveDocumentToFile(doc, "test.docx");
}

// Listen for clicks on Generate Word Document button
document.getElementById("generate").addEventListener(
  "click",
  function (event) {
    generateWordDocument(event);
  },
  false
);