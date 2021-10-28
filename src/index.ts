import { TableCreator } from './CreateTable';
import * as fs from "fs";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import { Schedule } from './Schedule';
//create static list of schedule elements
let weeklySchedule: Schedule[] = [
    {
        Name: 'Joe',
        Day: 'Monday',
        StartTime: '8am',
        EndTime: '12pm',
        Hours: 4
    },
    {
        Name: 'Bob',
        Day: 'Monday',
        StartTime: '10am',
        EndTime: '3pm',
        Hours: 5
    },
    {
        Name: 'Skeeter',
        Day: 'Tuesday',
        StartTime: '8am',
        EndTime: '5pm',
        Hours: 9
    },
    {
        Name: 'Scooter',
        Day: 'Tuesday',
        StartTime: '12pm',
        EndTime: '5pm',
        Hours: 5
    }
]
/**
 * Main function for this app
 */
function main(): void {
    //create a new instance of the table creator
    let tableCreator = new TableCreator();
    //create a new docx Document
    let doc = new Document({
        sections: [{
            children: [
                //Create the title
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Acme Co. Schedule",
                            bold: true,
                        }),

                    ],
                    heading: HeadingLevel.TITLE,
                }),
                //add week below
                new Paragraph({
                    children:[
                        new TextRun({
                            text: "Week of 11/1/2021",
                            bold:true,
                            italics:true
                        })
                    ],
                    heading: HeadingLevel.HEADING_3
                }),
                new Paragraph({}),
                //create message to user
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "This is the schedule for Acme co. Any changes should be requested with your supervisor",
                            italics: true
                        })
                    ]
                }),
                //create a table
                tableCreator.CreateTable(weeklySchedule)
            ]
        }

        ]
    });
    //save
    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("Schedule.docx", buffer);
    });
}

//call main function
main();