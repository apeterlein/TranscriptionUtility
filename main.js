"use strict";

$(document).ready(function () {
	$("#file").val("");
	let filename = "";
	const reader = new FileReader();
	const parser = new DOMParser();
	reader.onload = function (evt) {
		let zip, xml, lines, doc;
		try { zip = new PizZip(evt.target.result); }
		catch {
			logError("I couldn't extract the contents of that file", "Try again on a different browser", error);
			return;
		}
		try {
			xml = parser.parseFromString(zip.files["word/document.xml"].asText(), "application/xml");
			window.xml = xml;
			window.parser = parser;
			window.raw = zip.files["word/document.xml"].asText();
			const error = xml.querySelector("parsererror");
		}
		catch (error) {
			logError("MS Word has generated some yucky XML that I couldn't parse", "Try removing any links or headers from the document", error);
			return;
		}
		try {
			lines = parseInput(xml);
			doc = genOutput(lines);
		}
		catch (error) {
			logError("I ran into a problem formatting text", "Try removing any links or headers from the document", error);
			return;
		}
		try { saveOutput(doc, filename); }
		catch (error) { logError("I couldn't save the output file", "Try again on a different browser", error); }
	};
	$("#file").change(function () {
		$("#error").html("");
		$("#details").html("");
		const file = document.getElementById("file").files[0];
		if (file && /\.docx$/g.test(file.name)) {
			filename = file.name.substring(0, file.name.length - 5);
			try { reader.readAsArrayBuffer(file, "UTF-8"); }
			catch (error) { logError("I couldn't read content from that file", "Try again using a different browser", error); }
		}
		else { logError("only .docx files are supported", "Try saving as a .docx", "{ name: " + file.name + ", type: " + file.type + " }"); }
		$("#file").val("");
	});
	window.addEventListener("dragover", function (evt) { evt.preventDefault(); } );
	window.addEventListener("dragenter", function (evt) { evt.preventDefault(); } );
	window.addEventListener("drop", function (evt) {
		evt.preventDefault();
		document.getElementById("file").files = evt.dataTransfer.files;
		$("#file").trigger("change");
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
				const speaker = str.substring(9, str.length);
				if (!$("#cons").is(':checked') || !(lines[row] && speaker === lines[row]["Speaker"])) {
					row++;
					lines[row] = {};
					lines[row]["Time"] = str.substring(0, 8);
					lines[row]["Speaker"] = speaker;
					lines[row]["Text"] = [];
				}
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
			if (cont.length > 0) {
				if ($("#brks").is(':checked')) cont.push(new docx.TextRun({ break: 1 }));
				else cont.push(new docx.TextRun({ text: " " }));
			}
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
function logError(txt, hlp, error) {
	$("#error").html("Sorry, " + txt + " :(");
	$("#details").html(hlp + ". If you need help, send this junk to Adam: <span class=\"fw-light fst-italic\">" + error + "</span>");
}