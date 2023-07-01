"use strict";

$(document).ready(function () {
	$("#file").val("");
	let filename = "";
	const reader = new FileReader();
	reader.onload = function (evt) {
		const zip = new PizZip(evt.target.result);
		const xml = $.parseXML(zip.files["word/document.xml"].asText())
		const lines = parseInput(xml);
		const doc = genOutput(lines);
		saveOutput(doc, filename);
	};
	$("#file").change(function () {
		const file = document.getElementById("file").files[0];
		if (file && /\.docx$/g.test(file.name)) {
			filename = file.name.substring(0, file.name.length - 5);
			reader.readAsArrayBuffer(file, "UTF-8");
		}
		$("#file").val("");
	});
});
function parseInput(xml) {
	let str = "";
	let lines = [];
	let row = -1;
	for (let i = 0; i < xml.all.length; i++) {
		const tag = xml.all[i].tagName;
		if (tag === "w:p" || tag === "w:br") {
			str = str.trim();
			if (/^\d\d:\d\d:\d\d [a-zA-Z\d ]+/g.test(str)) {
				row++;
				lines[row] = {};
				lines[row]["Time"] = str.substring(0, 8);
				lines[row]["Speaker"] = str.substring(9, str.length);
				lines[row]["Text"] = [];
			}
			else if (str.length !== 0 && lines[row] !== undefined) {
				lines[row]["Text"].push(str);
			}
			str = "";
		}
		else if (tag === "w:t") {
			str += xml.all[i].innerHTML;
		}
	}
	return lines;
}
function genOutput(lines) {
	let cells = [
		new docx.TableRow({
			children: [
				new docx.TableCell({
					children: [
						new docx.Paragraph({
							children: [
								new docx.TextRun({
									text: "Timestamp",
									bold: true,
								})
							]
						})
					]
				}),
				new docx.TableCell({
					children: [
						new docx.Paragraph({
							children: [
								new docx.TextRun({
									text: "Speaker",
									bold: true,
								})
							]
						})
					]
				}),
				new docx.TableCell({
					children: [
						new docx.Paragraph({
							children: [
								new docx.TextRun({
									text: "Content",
									bold: true,
								})
							]
						})
					]
				})
			]
		})
	];
	for (let i = 0; i < lines.length; i++) {
		let cont = []
		for (let j = 0; j < lines[i]["Text"].length; j++) {
			cont.push(new docx.TextRun({ text: lines[i]["Text"][j] }));
		}
		let row = new docx.TableRow({
			children: [
				new docx.TableCell({
					children: [
						new docx.Paragraph({
							children: [
								new docx.TextRun({
									text: lines[i]["Time"]
								})
							]
						})
					]
				}),
				new docx.TableCell({
					children: [
						new docx.Paragraph({
							children: [
								new docx.TextRun({
									text: lines[i]["Speaker"]
								})
							]
						})
					]
				}),
				new docx.TableCell({
					children: [ new docx.Paragraph({ children: cont }) ]
				})
			]
		});
		cells.push(row);
	}
	return new docx.Document({
		sections: [{
			properties: {},
			children: [
				new docx.Table({
					width: { size: 9070, type: docx.WidthType.DXA },
					rows: cells
				})
			]
		}]
	});
}
function saveOutput(doc, filename) {
	docx.Packer.toBlob(doc).then(function (blob) { saveAs(blob, filename + "_out.docx"); });
}