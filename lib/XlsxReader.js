"use strict";

const EventEmitter = require('events');
const Transform = require('stream').Transform;
const StreamZip = require("node-stream-zip");
const Entities = require('html-entities').XmlEntities;
const xmlParseString = require('xml2js').parseString;
const _ = require('lodash');

/**
 * XlsxReader provides a very basic mechanism to fast-stream XLSX files.
 * @class XlsxReader
 * @extends {EventEmitter}
 * @fires XlsxReader#worksheet
 */
class XlsxReader extends EventEmitter {

    constructor(filePath, options) {
        super();

        this.options = Object.assign({
            ignoreEmpty: false,
            sheets: []
        }, options);

        let sheetTable = {};
        let stringTable = [];
        let waitingOnStrings = new EventEmitter();

        let zip = new StreamZip({
            file: filePath,
            storeEntries: true
        });

        let entities = new Entities();

        zip
            .on("error", (err) => {
                this.emit("error", err);
            })
            .on("entry", (entry) => {
                if (entry.name === "xl/sharedStrings.xml") {
                    zip.stream(entry.name, (err, stream) => {
                        let buffer = "";
                        let stringData = false;
                        stream
                            .on("data", (chunk) => {
                                buffer = `${buffer}${chunk.toString()}`;
                                if (stringData) {
                                    let strings = buffer.split("</t></si>");
                                    buffer = strings.pop();
                                    for (let index = 0; index < strings.length; index++) {
                                        let parts = strings[index].split(">");
                                        while (parts.length > 3) {
                                            stringTable.push("");
                                            parts = parts.slice(3);
                                        }

                                        stringTable.push(entities.decode(parts
                                            .pop()
                                            .trim()
                                        ));
                                    }
                                } else {
                                    let parts = buffer.split(/<sst.*?>/);
                                    buffer = parts.pop();
                                    stringData = (parts.length > 0);
                                }
                            })
                            .on("end", () => {
                                waitingOnStrings.emit("ready");
                            });
                    });
                }

                if (entry.name === "xl/_rels/workbook.xml.rels") {
                    const data = zip.entryDataSync(entry);
                    xmlParseString(data.toString(), (err, result) => {
                        result.Relationships.Relationship.forEach((relationship) => {
                            const str = 'worksheets';
                            if (relationship.$.Target.substring(0, str.length) === str) {
                                sheetTable[relationship.$.Id] = relationship.$.Target;
                            }
                        });
                    });
                }

                if (entry.name === "xl/workbook.xml") {
                    const data = zip.entryDataSync(entry);
                    xmlParseString(data.toString(), (err, result) => {
                        result.workbook.sheets[0].sheet.forEach((sheet) => {
                            if (this.options.sheets.length && this.options.sheets.indexOf(sheet.$.name) === -1) {
                                delete sheetTable[sheet.$['r:id']];
                            }
                        });
                    });
                }

                let matchEntry = (/xl\/(worksheets\/sheet(\d+)\.xml)$/).exec(entry.name);
                if (matchEntry && _.values(sheetTable).indexOf(matchEntry[1]) !== -1) {
                    let compressionRatio = 1.0;
                    /**
                      * @event XlsxReader#worksheet
                      * @type {object}
                      * @property {number} index - The (base-1) index of the sheet within the workbook.
                      * @function openReadStream - Open the sheet's stream
                      * @param {function} onProgress - Optional callback to report when raw bytes are read.
                      * @returns {Readable} - the stream of the sheet (rows are passed as arrays.)
                      */
                    this.emit("worksheet", {
                        index: Number(matchEntry[2]),
                        openReadStream: (onProgress) => {
                            let buffer = "";
                            let sheetData = false;
                            let lastRowNum = 0;
                            let sheet = new Transform({
                                objectMode: true,
                                transform: (chunk, encoding, next) => {
                                    if (onProgress) {
                                        onProgress(chunk.length / compressionRatio);
                                    }
                                    buffer = `${buffer}${chunk.toString()}`;
                                    if (!sheetData) {
                                        let parts = buffer.split("<sheetData>");
                                        buffer = parts.pop();
                                        sheetData = (parts.length > 0);
                                    }
                                    if (sheetData) {
                                        let rows = buffer.split("<row ");
                                        buffer = rows.pop();
                                        for (let row = 0; row < rows.length; row++) {
                                            if (!rows[row]) {
                                                continue;
                                            }
                                            const rowXml = `<row ${rows[row]}`;
                                            let output = [];
                                            // If we're not ignoring empty rows, we need to output them.
                                            if (!this.options.ignoreEmpty) {
                                                const rowNum = Number((/<row.*?r="(\d+)".*?>/).exec(rowXml)[1]);
                                                while (++lastRowNum < rowNum) {
                                                    sheet.push(output);
                                                }
                                            }

                                            xmlParseString(rowXml, (err, result) => {
                                                if (result.row.c) {
                                                    result.row.c.forEach((cell) => {
                                                        const column = cell.$.r.match(/([A-Z]{1,2})\d+?/);
                                                        if (column) {
                                                            let index = 0;
                                                            for (let i = 0; i < column[1].length; i++) {
                                                                index *= 26;
                                                                index += column[1].charCodeAt(i) - 64;
                                                            }
                                                            --index;
                                                            let value = cell.v;
                                                            if (value) {
                                                                if (cell.$.t === 's') {
                                                                    output[index] = stringTable[Number(value)];
                                                                } else {
                                                                    output[index] = Number(value);
                                                                }
                                                            }
                                                        }
                                                    });
                                                }
                                            });

                                            sheet.push(output);
                                        }
                                    }
                                    next();
                                },
                                flush: (done) => {
                                    this.emit("finish");
                                    done();
                                }
                            });

                            zip.stream(entry.name, (err, stream) => {
                                if (waitingOnStrings) {
                                    waitingOnStrings.on("ready", () => {
                                        compressionRatio = entry.size / entry.compressedSize;
                                        stream.pipe(sheet);
                                    });
                                } else {
                                    stream.pipe(sheet);
                                }
                            });
                            return sheet;
                        }
                    });
                }
            });
    }

    /**
     * @function objectify - Utility function to generate objects from row arrays.
     * @param {function} headers - Row array containg header keys.
     * @param {function} data - Row array containing row values.
     * @returns {object} - An object representing the union of the arrays.
     */
    objectify(headers, data) {
        let length =
            (data.length > headers.length) ?
            headers.length :
            data.length;

        let output = {};
        for (let i = 0; i < length; i++) {
            output[headers[i]] = data[i] || "";
        }
        return output;
    }
}

module.exports = XlsxReader;
