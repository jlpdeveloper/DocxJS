import { AlignmentType, TextRun, WidthType } from 'docx';
import { Paragraph, Table, TableCell, TableRow, VerticalAlign } from 'docx';
import { Schedule } from './Schedule';
export class TableCreator {
    public CreateTable(schedule: Schedule[]): Table {
        let rows = [];
        //create header row
        rows.push(new TableRow({
            children: [
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Day",
                                    bold: true
                                }),
                            ]
                        })
                    ]
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Employee",
                                    bold: true
                                }),
                            ]
                        })
                    ]
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Start Time",
                                    bold: true
                                }),
                            ]
                        })
                    ]
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "End Time",
                                    bold: true
                                }),
                            ],
                            
                        })
                    ]
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Hours",
                                    bold: true
                                }),
                            ]
                        })
                    ]
                })
            ]
        }));
        //loop thru each of the schedule elements
        schedule.forEach(s => {
            //push a new row element based on the schedule sent in
            rows.push(this._createRow(s));
        });
        //create a new table element and set to 100% width
        let table: Table = new Table({
            rows: rows,
            width:{
                size:100,
                type:WidthType.PERCENTAGE
            }
        });
        //return the table
        return table;
    }
    //function to create a table row easily
    private _createRow(schedule: Schedule): TableRow {
        //create a row with all the cells
        //NOTE: the children has to have a paragraph elements
        return new TableRow({
            children: [
                new TableCell({
                    children: [
                        new Paragraph({
                            text: schedule.Day,
                        }),
                    ]
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            text: schedule.Name,
                        }),
                    ]
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            text: schedule.StartTime,
                            alignment: AlignmentType.RIGHT
                        }),
                    ]
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            text: schedule.EndTime,
                            alignment: AlignmentType.RIGHT
                        }),
                    ]
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            text: schedule.Hours.toString(),
                            alignment: AlignmentType.RIGHT
                        }),
                    ]
                    
                })
            ]
        })
    }
};