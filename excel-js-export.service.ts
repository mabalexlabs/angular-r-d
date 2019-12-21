import { Injectable } from '@angular/core';

import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import { ExportService } from './export.service';
import { concat, filter } from 'rxjs/operators';

/**
 * @author Amiya Kumar Mahapatro
 * https://www.ngdevelop.tech/export-to-excel-in-angular-6/
 * 
 * npm install --save exceljs@1.12.0
 * 
 * update ts.config
 * "compilerOptions": {
    ...
    "paths": {
      "exceljs": [
        "node_modules/exceljs/dist/exceljs.min"
      ]
    }
  }
*
*  npm install --save file-saver
**/
@Injectable()
export class ExcelJsExportService {

  constructor(private exportService: ExportService) { }


  generateExcel(input): void {
    const finalData = this.formattingInput(input);
    this.generateExcel$1(finalData);
  }
  formattingInput(input: { reports: any, name: string, filter: any }) {
    const result: any = {};
    input.reports.forEach((report, i) => {
      const headersInfo = report.metaInfo.fields.filter(item => item.dataIndex !== 'seqNo');
      const metaInfo = headersInfo.map((item: any) => item.name);
      const metaInfoKey = headersInfo.map((item: any) => item.dataIndex);
      const gridData = [];
      let widthArr = [];
      let packets = 0;
      for (let index = 0; index < report.data.length; index++) {
        const dataStream = report.data.slice(packets, packets += 50);
        if (!dataStream.length) {
          break;
        }
        dataStream.forEach(object => {
          const temp = [];
          delete object.seqNo;
          headersInfo.forEach(key => {
            let value = object[key.dataIndex];
            if (value !== undefined && value !== null && key && 'integer'.includes(key.type)) {
              value = Number.parseInt(value, 10);
            } else if (value !== undefined && value !== null && key && 'float'.includes(key.type)) {
              value = Number.parseFloat(value);
            }
            value = value === undefined || value === null ? '' : value;
            temp.push(value);
          });
          gridData.push(temp);
        });

        if (!index) {
          widthArr = metaInfoKey.map((element, j) => {
            const longestGenre = Math.max(...dataStream.map(item => (item[element] || '').toString().length));
            return longestGenre > metaInfo[j].length ? longestGenre : metaInfo[j].length;
          });
        }
      }
      result['index-' + i] = {
        column_widths: widthArr,
        name: `Report Index - ${i + 1}`,
        headers: metaInfo,
        data: gridData,
        colInfo: headersInfo,
        report_name: input.name
      };

    });
    return result;
  }

  generateExcel$1(finalData) {
    const GLOBAL_HEADER_STYLE: any = { name: 'Arial', family: 4, size: 10, bold: true };
    const GLOBAL_CELL_STYLE: any = { name: 'Arial', family: 4, size: 10, bold: false };
    const GLOBAL_CELL_BORDER: any = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
    const workbook = new Workbook();
    //  first sheet created by deployment name
    const report_name = finalData['index-0'].report_name;
    const worksheet = workbook.addWorksheet(report_name);
    const filters = this.exportService.formatFilterData('reports', report_name);

    const headerstyle = (filter, isCenter) => {
      const temp = worksheet.addRow(filter);
      temp.font = GLOBAL_HEADER_STYLE;
      temp.border = {};
      temp.eachCell({ includeEmpty: false }, function (cell, colNumber) {
        cell.border = GLOBAL_CELL_BORDER;
        cell.alignment = { vertical: 'top', horizontal: isCenter ? 'center' : 'left' };
      });

    }
    filters.forEach((filter, i) => {
      headerstyle(filter, false);
    });

    const maxColumnLengths: Array<Array<number>> = []

    Object.keys(finalData).forEach((key, i) => {
      const gridData = finalData[key];
      worksheet.addRow([]);
      headerstyle(gridData.headers, true);
      // const rows = worksheet.addRows(gridData.data);
      gridData.data.forEach(record => {
        const row = worksheet.addRow(record);
        row.eachCell((cell, cellIndex) => {
          cell.border = GLOBAL_CELL_BORDER;
          cell.font = GLOBAL_CELL_STYLE;
          const colInfo = gridData.colInfo[cellIndex - 1];
          if (colInfo.type === 'integer') {
            cell.numFmt = '#,##0;[Red]-#,##0';
            cell.alignment = { vertical: 'top', horizontal: 'right' };
          } else if (colInfo.type === 'float') {
            cell.numFmt = '0.00';
            cell.alignment = { vertical: 'top', horizontal: 'right' };
          }
        });
      });

      maxColumnLengths.push(gridData.column_widths);

    });
    const longest = maxColumnLengths.reduce(function (a, b) { return a.length > b.length ? a : b; });
    const firstTH = maxColumnLengths[0];

    longest.forEach((length, l) => {
      if (l <= 1 && length < 40) {
        length = 40;
      } else if ((l + 1) <= firstTH.length) {
        length = firstTH[l];
      }
      worksheet.getColumn(l + 1).width = length + 5;
    });

    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const date = new Date();
      const fileNameWithDate = date.getFullYear() + '' +
        (date.getMonth() + 1) + '' + date.getDate() + '' +
        date.getHours() + '' + date.getMinutes() + '' +
        date.getSeconds();
      fs.saveAs(blob, report_name.split(' ').join('') + '_' + fileNameWithDate);
    });
  }


  // for multiple sheets
  generateExcel$2(finalData) {
    const workbook = new Workbook();
    const worksheets = [];
    Object.keys(finalData).forEach(key => {
      const gridData = finalData[key];
      worksheets.push(workbook.addWorksheet(gridData.name));
    });
    worksheets.forEach((ws, i) => {

      const filters = this.exportService.formatFilterData('reports', finalData[0].report_name);
      filters.forEach((filter, i) => {
        const temp = ws.addRow(filter.map(value => value.trim()));
        temp.font = { name: 'Arial', family: 4, size: 10, bold: true };
        temp.border = {};

        temp.eachCell({ includeEmpty: true }, function (cell, colNumber) {
          cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });
      });
      ws.addRow([]);
      const headerRow = ws.addRow(finalData['' + i].headers);
      headerRow.font = {
        name: 'Arial',
        family: 4, size: 10, bold: true
      };
      headerRow.eachCell((cell, number) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFF' },
          bgColor: { argb: 'FFFFFF' }
        };
        cell.alignment = { horizontal: 'center' };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });
      ws.addRows(finalData['' + i].data);
      finalData['' + i].headers.forEach((element, j) => {
        ws.getColumn(j + 1).eachCell((cell, rowNumber) => {
          if (!cell.font) {
            cell.font = {
              name: 'Arial',
              family: 4, size: 10
            };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
          }
        });
      });



      finalData['' + i].column_widths.forEach((width: number, l) => {
        const datatype = finalData['' + i].colInfo[l];
        ws.getColumn(l + 1).width = width + 5;
        if ([0, 1].includes(l)) {
          ws.getColumn(l + 1).width = 45;
        }
        if (datatype.type === 'integer') {
          ws.getColumn(l + 1).numFmt = '#,##0;[Red]-#,##0';

        } else if (datatype.type === 'float') {
          ws.getColumn(l + 1).numFmt = '0.00';
        } else {
          ws.getColumn(l + 1).eachCell((cell, rowNumber) => {
            cell.alignment = { vertical: 'top', horizontal: 'left' }
          });
        }
      });
    });



    // Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const date = new Date();
      const fileNameWithDate = date.getFullYear() + '' +
        (date.getMonth() + 1) + '' + date.getDate() + '' +
        date.getHours() + '' + date.getMinutes() + '' +
        date.getSeconds();
      fs.saveAs(blob, finalData[0].report_name.split(' ').join('') + '_' + fileNameWithDate);
    });
  }

}
