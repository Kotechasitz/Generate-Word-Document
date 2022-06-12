import { Document, Packer, Paragraph, HeadingLevel, Table, TableCell, TableRow, WidthType, Alignment, AlignmentType, ShadingType } from "docx";
import { saveAs } from "file-saver";
import * as fs from "fs";

let warningTable = new Table({
    columnWidths: [7208, 2000],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.DXA,
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
                        type: WidthType.DXA,
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
                        type: WidthType.DXA,
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
                        type: WidthType.DXA,
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
                        type: WidthType.DXA,
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
                        type: WidthType.DXA,
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
                        type: WidthType.DXA,
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
    ],
});

let selfControlChecklistTable = new Table({
    columnWidths: [1000,6208, 2000],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "No", style: "textBlackColor", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "92d050",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"Details", style: "textBlackColor",alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "92d050",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "*Yes/No", style: "textRedColor", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "92d050",
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
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "1", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"Check log, alarm และ service ก่อนทำการ deploy application ว่าสามารถทำงานได้ตามปกติ", style: "textBlackColor"})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Yes", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
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
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "2", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"Check ข้อมูลใน WI",style: "textNomal",}),
                    new Paragraph({text:"- Software version",style: "textNomal",}),
                    new Paragraph({text:"- Configuration",style: "textNomal",}),
                    ],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"",style: "textNomal",}),
                    new Paragraph({text:"Yes",style: "textNomal", alignment: AlignmentType.CENTER}),
                    new Paragraph({text:"",style: "textNomal",}),],
                    shading: {
                        fill: "ccff99",
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
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "3", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"ทำการ Backup file และ configuration", style: "textNomal"})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Yes", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
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
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "4", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:`Copy source code จากเครื่อง staging มาไว้บนเครื่อง production
                    ทำการตรวจสอบ source code อีกครั้ง`, style: "textNomal"})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Yes", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
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
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "5", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"ทำการ Deploy application", style: "textNomal"}),
                    new Paragraph({text:"- Edit configuration", style: "textNomal"}),
                    new Paragraph({text:"- Compile code", style: "textNomal"}),
                    new Paragraph({text:"- Reload application", style: "textNomal"})
                    ],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"", style: "textNomal"}),
                    new Paragraph({text:"", style: "textNomal"}),
                    new Paragraph({text:"Yes", style: "textNomal",alignment: AlignmentType.CENTER}),
                    new Paragraph({text:"", style: "textNomal"})],
                    shading: {
                        fill: "ccff99",
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
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "6", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"Check log, alarm และ service หลังทำการ deploy application ว่าสามารถทำงานได้ตามปกติ", style: "textRedColor",alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph("")],
                    shading: {
                        fill: "ccff99",
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
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "7", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7208,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"แจ้งผลการ Deploy เพื่อให้ทาง Solutions ทำการ Post Test ต่อไป", style: "textNomal"})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Yes", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
    ],
});

let detailsprogramOnProductionTable = new Table({
    columnWidths: [2500,3200, 3508],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 2500,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "SCR/UR", style: "textBlackColor", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fef010",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 3200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "SIR", style: "textBlackColor", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fef010",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 3508,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Description", style: "textBlackColor", alignment: AlignmentType.CENTER})],
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
                        size: 2500,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-069945", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "fef010",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 3200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fef010",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 3508,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "SP22.5.3 Deploy Phoenix New Inventory (VHL) on 20220601", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fef010",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
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
        },
        {
            id: "textBlackColor",
            name: "Text Black Color",
            basedOn: "Normal",
            run: {
              color: "030303",
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
          warningTable,
          new Paragraph(""),
          new Paragraph({ text: "Self Control Checklist (for Operations)",
          style: "textHeader",
          alignment: AlignmentType.CENTER
          }),
          selfControlChecklistTable,
          new Paragraph(""),
          new Paragraph({ text: "1. Details Program on Production",
          style: "textHeader",
          }),
          detailsprogramOnProductionTable,
          new Paragraph(""),
          new Paragraph({ text: "2. Impact",
          style: "textHeader",
          }),
        ],
      }
    ],
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