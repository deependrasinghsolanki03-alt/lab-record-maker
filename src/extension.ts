import * as vscode from 'vscode';
import * as cp from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { 
    Document, 
    Packer, 
    Paragraph, 
    TextRun, 
    AlignmentType, 
    HeadingLevel, 
    PageBreak, 
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    ShadingType
} from 'docx';

interface StudentProfile {
    name: string;
    enrollment: string;
    course: string;
}

function sanitizeText(text: string): string {
    if (!text) return "";
    return text.replace(/[^\x20-\x7E\n\t]/g, '');
}

function runCode(filePath: string, inputValue: string): Promise<string> {
    return new Promise((resolve) => {
        const outPath = filePath.replace('.cpp', '.exe'); 
        const child = cp.exec(`g++ "${filePath}" -o "${outPath}" && "${outPath}"`, { timeout: 5000 }, (err, stdout, stderr) => {
            if (err) resolve(`ERROR:\n${sanitizeText(stderr || 'Execution failed')}`);
            else resolve(sanitizeText(stdout));
        });

        if (child.stdin) {
            if (inputValue) {
                const cleanInput = sanitizeText(inputValue).replace(/[ ,]+/g, '\n'); 
                child.stdin.write(cleanInput + '\n');
            }
            child.stdin.end();
        }
    });
}

export function activate(context: vscode.ExtensionContext) {

    async function getStudentProfile(forceEdit = false): Promise<StudentProfile | undefined> {
        let profile = context.globalState.get<StudentProfile>('studentProfile');
        if (!profile || forceEdit) {
            const name = await vscode.window.showInputBox({ prompt: "Student Name", value: profile?.name || "" });
            if (!name) return undefined;
            const enrollment = await vscode.window.showInputBox({ prompt: "Enrollment No.", value: profile?.enrollment || "" });
            if (!enrollment) return undefined;
            const course = await vscode.window.showInputBox({ prompt: "Course Name", value: profile?.course || "" });
            if (!course) return undefined;

            profile = { name, enrollment, course };
            await context.globalState.update('studentProfile', profile);
        }
        return profile;
    }

    let editCommand = vscode.commands.registerCommand('extension.editStudentDetails', async () => {
        await getStudentProfile(true);
    });

    let mainCommand = vscode.commands.registerCommand('extension.appendCodeToRecord', async () => {
        
        const editor = vscode.window.activeTextEditor;
        if (!editor) return vscode.window.showErrorMessage("Open a .cpp file first!");

        await editor.document.save();
        const profile = await getStudentProfile(false); 
        if (!profile) return;

        const rawCode = editor.document.getText();
        const cleanCode = sanitizeText(rawCode);

        const rootPath = vscode.workspace.workspaceFolders?.[0].uri.fsPath;
        if (!rootPath) return vscode.window.showErrorMessage("Open a Folder first!");

        let docName = await vscode.window.showInputBox({
            placeHolder: "Lab_Record",
            prompt: "Word File Name (Bina .docx ke)",
            value: "Lab_Record"
        });
        if (!docName) return;
        if (!docName.endsWith('.docx')) docName += '.docx';

        const userInput = await vscode.window.showInputBox({
            placeHolder: "Inputs (e.g., 5 10)",
            prompt: "Inputs dein (Optional)",
        });

        vscode.window.withProgress({
            location: vscode.ProgressLocation.Notification,
            title: `Creating Professional Doc: ${docName}...`,
            cancellable: false
        }, async () => {
            
            const currentOutput = await runCode(editor.document.fileName, userInput || "");

            const historyFilename = docName.replace('.docx', '_history.json');
            const historyPath = path.join(rootPath, historyFilename);
            const docPath = path.join(rootPath, docName);

            let experiments = [];
            if (fs.existsSync(historyPath)) {
                try { experiments = JSON.parse(fs.readFileSync(historyPath, 'utf-8')); } 
                catch(e) { experiments = []; }
            }

            experiments.push({
                expNo: experiments.length + 1,
                code: cleanCode,
                inputUsed: userInput ? sanitizeText(userInput) : "",
                output: currentOutput,
                date: new Date().toLocaleString()
            });

            fs.writeFileSync(historyPath, JSON.stringify(experiments, null, 2));

            // DOC GENERATION
            const experimentChildren: any[] = [];

            for (const exp of experiments) {
                const lines = exp.code.split('\n');
                let aim = "Aim not provided";
                let codeLines = lines; 

                if (lines.length > 0 && lines[0].trim().startsWith('//')) {
                    aim = lines[0].trim().substring(2).trim();
                    codeLines = lines.slice(1);
                }

                // --- 1. CODE BOX (TABLE METHOD - FIXED) ---
                // Yahan 'line: string' likhne se error hat jayega
                const codeParagraphs = codeLines.map((line: string) => 
                    new Paragraph({
                        children: [new TextRun({ text: line, font: "Consolas", size: 20 })],
                        spacing: { after: 0, before: 0 }
                    })
                );

                const codeTable = new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: codeParagraphs,
                                    shading: { fill: "F2F2F2", type: ShadingType.CLEAR },
                                    borders: {
                                        top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
                                        bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
                                        left: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
                                        right: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
                                    },
                                    margins: { top: 100, bottom: 100, left: 100, right: 100 }
                                })
                            ]
                        })
                    ]
                });

                // --- 2. ADDING TO DOC ---
                experimentChildren.push(
                    new Paragraph({
                        text: `Experiment ${exp.expNo}`,
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 400, after: 100 }
                    }),
                    new Paragraph({
                        text: `Date: ${exp.date}`,
                        alignment: AlignmentType.RIGHT,
                        spacing: { after: 300 }
                    }),
                    new Paragraph({
                        children: [new TextRun({ text: "Aim/Question:", bold: true, font: "Times New Roman", size: 28 })],
                        spacing: { after: 100 }
                    }),
                    new Paragraph({
                        children: [new TextRun({ text: aim, font: "Times New Roman", size: 24 })],
                        spacing: { after: 300 }
                    }),
                    new Paragraph({
                        children: [new TextRun({ text: "Source Code:", bold: true, font: "Times New Roman", size: 28 })],
                        spacing: { after: 100 }
                    }),
                    
                    codeTable, // CODE BOX TABLE

                    new Paragraph({ text: "", spacing: { after: 300 } })
                );

                if (exp.inputUsed) {
                    experimentChildren.push(
                        new Paragraph({
                            children: [new TextRun({ text: "Input Provided:", bold: true, color: "FF0000", font: "Times New Roman" })],
                        }),
                        new Paragraph({
                            children: [new TextRun({ text: exp.inputUsed, font: "Consolas", size: 20 })],
                            spacing: { after: 200 }
                        })
                    );
                }

                const outputParagraphs = exp.output.split('\n').map((line: string) => 
                    new Paragraph({
                        children: [new TextRun({ text: line, font: "Consolas", size: 20, color: "0000FF" })],
                        spacing: { after: 0 }
                    })
                );

                experimentChildren.push(
                    new Paragraph({
                        children: [new TextRun({ text: "Output:", bold: true, font: "Times New Roman", size: 28 })],
                        spacing: { after: 100 }
                    }),
                    ...outputParagraphs,
                    new Paragraph({
                        children: [new PageBreak()],
                        spacing: { before: 400 }
                    })
                );
            }

            const doc = new Document({
                sections: [{
                    children: [
                        new Paragraph({
                            text: "LAB PRACTICAL RECORD",
                            heading: HeadingLevel.TITLE,
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 500 }
                        }),
                        new Paragraph({
                            children: [new TextRun({ text: `Name: ${profile.name}`, font: "Times New Roman", size: 28 })],
                            alignment: AlignmentType.CENTER,
                        }),
                        new Paragraph({
                            children: [new TextRun({ text: `Enrollment: ${profile.enrollment}`, font: "Times New Roman", size: 28 })],
                            alignment: AlignmentType.CENTER,
                        }),
                        new Paragraph({
                            children: [new TextRun({ text: `Course: ${profile.course}`, font: "Times New Roman", size: 28 })],
                            alignment: AlignmentType.CENTER,
                            spacing: { after: 800 }
                        }),
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                                new TextRun("-----------------------------------------------------------"),
                                new PageBreak() 
                            ]
                        }),
                        ...experimentChildren
                    ],
                }],
            });

            const buffer = await Packer.toBuffer(doc);
            fs.writeFileSync(docPath, buffer);

            vscode.window.showInformationMessage(`Professional Doc Created: ${docName}`);
        });
    });

    context.subscriptions.push(mainCommand);
    context.subscriptions.push(editCommand);
}

export function deactivate() {}