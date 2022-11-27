import { Document, Packer, Paragraph, HeadingLevel, Table, TableCell, TableRow, WidthType, Alignment, AlignmentType, ShadingType } from "docx";
import { saveAs } from "file-saver";
import * as fs from "fs";

let warningTable = new Table({
    columnWidths: [500,7708, 1000],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 500,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "#", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffff00",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7708,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"Details", style: "textBlackColorBold",alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffff00",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Check", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffff00",
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
                        size: 500,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "1", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fff2cc",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7708,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:" Check log, alarm และ service ก่อนทำการ deploy application ว่าสามารถทำงานได้ตามปกติ", style: "textBlackColor"})],
                    shading: {
                        fill: "fff2cc",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Yes", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fff2cc",
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
                        size: 500,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "2", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "fff2cc",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:" Check ข้อมูลใน WI: Software version, Checksum, Configuration",style: "textNomal",}),
                    ],
                    shading: {
                        fill: "fff2cc",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"Yes",style: "textNomal", alignment: AlignmentType.CENTER})
                    ],
                    shading: {
                        fill: "fff2cc",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
    ],
});

let selfControlChecklistTable = new Table({
    columnWidths: [500,7708, 1000],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 500,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "#", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "92d050",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 7708,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"Details", style: "textBlackColorBold",alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "92d050",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Check", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
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
                        size: 500,
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
                        size: 7708,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:" Check log, alarm และ service ก่อนทำการ deploy application ว่าสามารถทำงานได้ตามปกติ", style: "textBlackColor"})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
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
                        size: 500,
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
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:" Check ข้อมูลใน WI: Software version, Checksum, Configuration",style: "textNomal",}),
                    ],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"Yes",style: "textNomal", alignment: AlignmentType.CENTER})
                    ],
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
                        size: 500,
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
                        size: 7708,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:" ทำการ Backup file และ configuration", style: "textNomal"})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
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
                        size: 500,
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
                        size: 7708,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:` Copy source code จากเครื่อง staging มาไว้บนเครื่อง production และทำการตรวจสอบ source code อีกครั้ง`, style: "textNomal"})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
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
                        size: 500,
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
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:" ทำการ Deploy application: Edit configuration, Compile code, Reload application", style: "textNomal"}),
                    ],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:"Yes", style: "textNomal",alignment: AlignmentType.CENTER}),
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
                        size: 500,
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
                        size: 7708,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:" Check log, alarm และ service หลังทำการ deploy application ว่าสามารถทำงานได้ตามปกติ", style: "textNomal",alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
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
                        size: 500,
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
                        size: 7708,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({text:" แจ้งผลการ Deploy เพื่อให้ทาง Solutions ทำการ Post Test ต่อไป", style: "textNomal"})],
                    shading: {
                        fill: "ccff99",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1000,
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
    columnWidths: [2200,2200, 4808],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "SCR/UR", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "SIR", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Description", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
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
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-069945", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "SP22.5.3 Deploy Phoenix New Inventory (VHL) on 20220601", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
    ],
});

let destinaTionSystemTable = new Table({
    columnWidths: [1841,1841,1841,1841,1841],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 1841,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Sever Name", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1841,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Server IP", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1841,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Database", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1841,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Domain", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1841,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
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
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "PNEWINVW801G", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "10.15.35.169", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-110378", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
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
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "PNEWINVW802G", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "10.15.35.200", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-110378", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
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
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "PNEWINVA801G", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "10.197.72.175", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-110378", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
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
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "PNEWINVA802G", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "10.197.72.177", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-110378", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
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
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "PNEWINVA803G", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "10.197.79.87", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-110378", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
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
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "PNEWINVD801G", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2200,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "10.13.140.49", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "MongoDB", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "-", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 4808,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-110378", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
    ],
});

let deploymentInstructionTable = new Table({
    columnWidths: [500, 1432, 1432, 1432, 1432, 1432, 1432],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "#", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "PCR/SIR", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Nexus", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "File Name", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Server Name", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Deploy Date", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Owner", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
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
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "1", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-110378", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "newinvent-releases/app/phxinvapp", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "phxinvapp.zip", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "PNEWINVA801G,PNEWINVA802G,PNEWINVA803G", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "24/08/202222.00", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1315,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Kotechasit Nilnont", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
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
                        size: 500,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "1", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1432,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "WR22-110378", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1432,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "newinvent-releases/app/phxinvapp", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1432,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "phxinvapp.zip", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1432,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "PNEWINVA801G,PNEWINVA802G,PNEWINVA803G", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1432,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "24/08/202222.00", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 1432,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Kotechasit Nilnont", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
    ],
});

let stepDeployTableOne = new Table({
    columnWidths: [9208],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 9208,
                        type: WidthType.DXA,
                    },
                    children: [
                        new Paragraph({ text: "$ cd /app", style: "textNomal", alignment: AlignmentType.LEFT}),
                        new Paragraph({ text: "$ mv phxinvapp “phxinvapp@$(date '+%Y%m%d')”", style: "textNomal", alignment: AlignmentType.LEFT}),
                        new Paragraph({ text: "$ unzip phxinvapp.zip", style: "textNomal", alignment: AlignmentType.LEFT}),
                        new Paragraph({ text: "$ sh /app/phx-conf/script/pm2-start.sh", style: "textNomal", alignment: AlignmentType.CENLEFTTER})
                    ],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
    ],
});

let stepDeployTableTwo = new Table({
    columnWidths: [9208],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 9208,
                        type: WidthType.DXA,
                    },
                    children: [
                        new Paragraph({ text: "$ cd /app", style: "textNomal", alignment: AlignmentType.LEFT}),
                        new Paragraph({ text: "$ mv phxinvweb “phxinvweb@$(date '+%Y%m%d')”", style: "textNomal", alignment: AlignmentType.LEFT}),
                        new Paragraph({ text: "$ unzip phxinvweb.zip", style: "textNomal", alignment: AlignmentType.LEFT}),
                        new Paragraph({ text: "$ sh /app/phx-conf/script/pm2-start.sh", style: "textNomal", alignment: AlignmentType.CENLEFTTER})
                    ],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
    ],
});

let updateDetailTable = new Table({
    columnWidths: [2302,6906],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 2302,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Main Feature", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 6906,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Sub Feature", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
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
                        size: 2302,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 6906,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "", style: "textNomal", alignment: AlignmentType.LEFT})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
            ],
        }),
    ],
});

let contractPersonTable = new Table({
    columnWidths: [500,2903,2903,2903],
    rows: [
        new TableRow({
            children: [
                new TableCell({
                    width: {
                        size: 500,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "#", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2903,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Name", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2903,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "E-mail", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2903,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "Phone Number", style: "textBlackColorBold", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ccff66",
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
                        size: 500,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "1", style: "textNomal", alignment: AlignmentType.CENTER}),],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2903,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2903,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
                        type: ShadingType.CLEAR,
                        color: "auto",
                    },
                }),
                new TableCell({
                    width: {
                        size: 2903,
                        type: WidthType.DXA,
                    },
                    children: [new Paragraph({ text: "", style: "textNomal", alignment: AlignmentType.CENTER})],
                    shading: {
                        fill: "ffffff",
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

function generateWordDocument() {
//   event.preventDefault();
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
              size: 28,
              font: "Times New Roman"
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
        },
        {
            id: "textBlackColorBold",
            name: "Text Black Color",
            basedOn: "Normal",
            run: {
              color: "030303",
              size: 22,
              font: "Sarabun",
              bold: true
            }
        },
        {
            id: "textSmall",
            name: "Text Small",
            basedOn: "Normal",
            run: {
              color: "999999",
              size: 18,
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
          new Paragraph({ text: "/* ผลกระทบจากการ change/deploy เช่น เกิด 01.00 AM หรือทำให้ระบบงานใดใช้งานไม่ได้ Function งานใดใช้งานไม่ได้บ้าง เป็นต้น */",
          style: "textSmall",
          }),
          new Paragraph(""),
          new Paragraph({ text: "3. Destination System",
          style: "textHeader",
          }),
          destinaTionSystemTable,
          new Paragraph(""),
          new Paragraph({ text: "4. Deployment Instruction",
          style: "textHeader",
          }),
          new Paragraph(""),
          new Paragraph({ text: "4.1 Web Application",
          style: "textSmall",
          }),
          new Paragraph({ text: "WR22-110378 SP22.8.2 Deploy Phoenix New Inventory (VHL) on 20220825",
          style: "textSmall",
          }),
          deploymentInstructionTable,
          new Paragraph(""),
          new Paragraph({ text: "4.2 Step การ Deploy",
          style: "textHeader",
          }),
          new Paragraph({ text: "Deploy ผ่านระบบ Nexus Jenkins CI/CD",
          style: "textSmall",
          }),
          new Paragraph(""),
          new Paragraph({ text: "PHXINVAPP(PNEWINVA801G, PNEWINVA802G, PNEWINVA803G)",
          style: "textNomal",
          }),
          stepDeployTableOne,
          new Paragraph(""),
          new Paragraph({ text: "PHXINVWEB(PNEWINVW801G, PNEWINVW802G)",
          style: "textNomal",
          }),
          new Paragraph(""),
          stepDeployTableTwo,
          new Paragraph(""),
          new Paragraph({ text: "5. Update Detail",
          style: "textHeader",
          }),
          updateDetailTable,
          new Paragraph(""),
          new Paragraph({ text: "6. Contact Person",
          style: "textHeader",
          }),
          contractPersonTable,
          new Paragraph(""),
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