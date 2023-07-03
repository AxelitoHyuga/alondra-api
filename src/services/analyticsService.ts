// import mysql from "mysql";
import { getCustomerAccountReceivables, getCustomerPaymentReconcilied } from "../models/analyticsModel";
import { dateStringFormatSql, getExcelColumnLetter } from "../tools";
import { AnalyticsFilters, Interval, RowAccountReceivable } from "../types";
import * as ExcelJS from "exceljs";

/* --- Estilos --- */
const HORIZONTAL_CENTER_ALIGNMENT: Partial<ExcelJS.Alignment> = {
    horizontal: 'center'
};
const HORIZONTAL_RIGTH_ALIGNMENT: Partial<ExcelJS.Alignment> = {
    horizontal: 'right'
};
const CURRENCY_FORMAT = '"$"#,##0.00;[Red]\-"$"#,##0.00';
const STYLE_FILTER_TEXT_TITLE: Partial<ExcelJS.Style> = {
    font: {
        bold: true,
        color: { argb: '00676D' }
    },
    alignment: {
        horizontal: 'left',
        vertical: 'middle',
    },
    fill: {
        type: 'pattern',
        pattern:'solid',
        fgColor: { argb: 'FFE6FFF2' }
    }
};
const STYLE_FILTER_TEXT: Partial<ExcelJS.Style> = {
    font: {
        italic: true,
        color: { argb: '00676D' }
    },
    alignment: {
        horizontal: 'left',
        vertical: 'top',
        wrapText: true,
        indent: 1
    },
    fill: {
        type: 'pattern',
        pattern:'solid',
        fgColor: { argb: 'FFE6FFF2' }
    }
};
const STYLE_HEAD: Partial<ExcelJS.Style> = {
    font: {
        bold: true,
        color: { argb: 'FFFFFF' }
    },
    alignment: {
        horizontal: 'center',
        vertical: 'middle'
    },
    border: {
        bottom: {
            style: 'thick',
            color: { argb: 'FFD240' }
        }
    },
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00676D' }
    }
};
const STYLE_ROW_STRIPED: Partial<ExcelJS.Style> = {
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF3F3F3' }
    }
};
const STYLE_FOOT: Partial<ExcelJS.Style> = {
    font: {
        bold: true
    },
    alignment: {
        vertical: 'middle'
    },
    border: {
        top: {
            style: 'thin',
            color: { argb: 'FF666666' }
        },
        bottom: {
            style: 'thin',
            color: { argb: 'FF666666' }
        }
    },
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFCCCCCC' }
    }
};
const STYLE_SOFT_VERSION: Partial<ExcelJS.Style> = {
    font: {
        italic: true,
        color: { argb: '888888' },
        size: 8
    }
};

/**
 * Consulta las cuentas por cobrar del sistema, y genera un archivo excel de los registros obtenidos.
 * - Nota: Para optimizar los tiempos de ejecución, se utilizan los métodos de iteración de 
 * {@link Array} ({@link Array.forEach}, {@link Array.map}, etc)
 * en lugar de los ciclos convencionales (for, while, etc).
 * @param filters Filtros para la consulta {@link AnalyticsFilters}
 * @returns El objeto xlsx del libro de trabajo {@link ExcelJS.Xlsx}
 */
const generateAccountReceivableExcel = async (filters: AnalyticsFilters): Promise<ExcelJS.Xlsx> => {
    return new Promise(async(resolve, _reject) => {
        const intervals: Interval[] = createDateIntervalPerWeek(filters.dateTo);
        const customerInvoices: RowAccountReceivable[] = await getCustomerAccountReceivables(filters, intervals);
        
        /* --- Se crea un nuevo libro de Excel --- */
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Sheet 1');

        /* --- Inicio --- */
        let rowIni = 1;
        const colEndNumber = 9 + intervals.length;
        const colEnd = getExcelColumnLetter(colEndNumber);
        let cell = worksheet.getCell(`A${rowIni}`);
        worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
        cell.value = 'Cuentas por cobrar';
        cell.style = STYLE_FILTER_TEXT_TITLE;

        /* --- Textos de Filtros --- */
        let filterText = '';
        filterText +=  filters.name != null ? `${filterText.length > 0 ? '\n' : ''} Número: ${filters.name}` : '';
        filterText +=  filters.dateFrom != null || filters.dateTo ? `${filterText.length > 0 ? '\n' : ''} Fecha: Desde ${filters.dateFrom != null ? filters.dateFrom : '--'} hasta ${filters.dateTo != null ? filters.dateTo : '--'}` : '';
        filterText +=  filters.origin != null ? `${filterText.length > 0 ? '\n' : ''} Doc. Origen: ${filters.origin}` : '';
        filterText +=  filters.customer != null ? `${filterText.length > 0 ? '\n' : ''} Cliente: ${filters.customer}` : '';
        filterText +=  filters.salesperson != null ? `${filterText.length > 0 ? '\n' : ''} Vendedor: ${filters.salesperson}` : '';
        filterText +=  `${filterText.length > 0 ? '\n' : ''} Fecha de creación: ${new Date().toLocaleDateString('es-mx')}`;

        /* --- Encabezado --- */
        rowIni++;
        cell = worksheet.getCell(`A${rowIni}`);
        worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
        cell.value = filterText;
        cell.style = STYLE_FILTER_TEXT;
        worksheet.getRow(rowIni).height = (filterText.split('\n').length * 13) + 6;

        /* --- Ancho de columnas --- */
        worksheet.getColumn('A').width = 35;
        worksheet.getColumn('B').width = 20;
        worksheet.getColumn('C').width = 12;
        worksheet.getColumn('D').width = 12;
        worksheet.getColumn('E').width = 30;
        worksheet.getColumn('F').width = 30;
        worksheet.getColumn('G').width = 20;
        worksheet.getColumn('H').width = 15;
        worksheet.getColumn(colEnd).width = 25;

        /* --- Títulos de las columnas --- */
        const titleRowData: string[] = [
            'Cliente',
            'Fecha',
            'Número',
            'Doc. Origen',
            'Referencia',
            'Vendedor',
            'Fecha de Vencimiento',
            'Moneda'
        ];
        
        const intervalColLetter: string[] = [];
        intervals.forEach((interval, key) => {
            const colLetter = getExcelColumnLetter(9 + key);
            titleRowData.push(interval.title);
            intervalColLetter[key] = colLetter;
            worksheet.getColumn(colLetter).width = 25;
        });

        titleRowData.push('Saldo Total');
        const titleRow = worksheet.addRow(titleRowData);
        titleRow.height = 30;
        titleRow.eachCell((cell) => {
            cell.style = STYLE_HEAD;
        });
        rowIni++;

        const dateTo = new Date(dateStringFormatSql(filters.dateTo));

        /* --- Registros --- */
        let row = (rowIni);
        console.time('customerInvoices');
        const promises = customerInvoices.map(async(result) => {
        // for (const result of customerInvoices) {
            const resultsPayment = await getCustomerPaymentReconcilied(result.customer_id);
            let balancePayment = result.total_balance;
            let totalAmountPayment = 0;
            let customerInvoicesId = new Map<string, boolean>(result.customer_invoice_id.split(',').map(v => [v , true]));
            let customerInvoicesId2 = new Map<string, boolean>(result.customer_invoice_id.split(',').map(v => [v , true]));
            let customerInvoicesId3 = new Map<string, boolean>(result.customer_invoice_id.split(',').map(v => [v , true]));

            if (filters.dateTo != null) {
                resultsPayment.forEach((value) => {
                    if (customerInvoicesId2.has(''+value.customer_invoice_id)) {
                        customerInvoicesId2.delete(''+value.customer_invoice_id);
                    } else if (customerInvoicesId.has(''+value.customer_invoice_id)) {
                        customerInvoicesId.set(''+value.customer_invoice_id, true);
                    }

                    if (new Date(value.customer_credit_date_credit) <= dateTo && customerInvoicesId.has(''+value.customer_invoice_id)) {
                        customerInvoicesId.delete(''+value.customer_invoice_id);
                        totalAmountPayment += value.amount;
                    } else if (new Date(value.customer_credit_date_credit) > dateTo && customerInvoicesId.has(''+value.customer_invoice_id)) {
                        customerInvoicesId.set(''+value.customer_invoice_id, true);
                        balancePayment += value.amount;
                    }
                });
            }

            // totalBalance += balancePayment;
            if (balancePayment <= 0) {
                return null;
            }

            /* --- Valores de Columnas --- */
            const bodyRowData: (string | number | Date)[] = [
                result.customer,
                result.date_invoice,
                result.name,
                result.origin,
                result.reference,
                result.salesperson,
                result.date_invoice,
                'MXN'
            ];

            /* Intervalos */
            // const _rowsRight: string[] = 
            intervals.map((interval, key) => {
                const colLetter = intervalColLetter[key];
                const tmpName = interval.name;
                let val = result[tmpName as keyof RowAccountReceivable] as number;
                let tmpCost = 0;

                if(filters.dateTo){
                    resultsPayment.forEach(value => {
                        if (tmpName == 'past_due') {
                            if (new Date(value.customer_credit_date_credit) <= dateTo && customerInvoicesId3.has(''+value.customer_invoice_id)) {
                                if(totalAmountPayment > 0 && val > tmpCost + value.amount){
                                    tmpCost += value.amount;
                                }
                            }
                        } else {
                            if (new Date(value.customer_credit_date_credit) < dateTo && customerInvoicesId3.has(''+value.customer_invoice_id)) {
                                if(totalAmountPayment > 0 && val > tmpCost + value.amount){
                                    tmpCost += value.amount;
                                }
                            }    
                        }
                    });
                }

                if (val > 0) {
                    val -= tmpCost;
                }

                bodyRowData.push(val);
                return colLetter;
            });
            
            bodyRowData.push(balancePayment);
            return bodyRowData;
        });
        
        Promise.all(promises).then((rows) => {
            console.timeEnd('customerInvoices');
            const rowsU = rows.filter(row => row);
            if (rowsU.length > 0) {
                const createdRows = worksheet.addRows(rowsU);
                createdRows.forEach(row => {
                    /* Estilo de celdas */
                    row.eachCell((cell) => {
                        const col = +cell.col;
                        if ([2, 3, 4, 5, 7, 8].includes(col)) {
                            cell.alignment = HORIZONTAL_CENTER_ALIGNMENT;
                        } else if (col > 8 && col < colEndNumber) {
                            cell.alignment = HORIZONTAL_RIGTH_ALIGNMENT;
                            cell.numFmt = CURRENCY_FORMAT;
                        }

                        if (col === 1) {
                            cell.border = {
                                left: {
                                    style: 'thin',
                                    color: { argb: 'FF666666' }
                                }
                            };
                        } else if (col === colEndNumber) {
                            cell.border = {
                                right: {
                                    style: 'thin',
                                    color: { argb: 'FF666666' }
                                }
                            };
                            cell.numFmt = CURRENCY_FORMAT;
                        }

                        if (+cell.row % 2) {
                            cell.style = Object.assign(cell.style, STYLE_ROW_STRIPED);
                        }
                    });
                });
            }
            /* --- Totales --- */
            const totalRowData: (string | ExcelJS.CellValue)[] = [
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                ''
            ];
            /* Intervalos */
            intervals.forEach((_interval, key) => {
                const colLetter = intervalColLetter[key];
                totalRowData.push({
                    formula: `=SUM(${colLetter}${(rowIni + 1)}:${colLetter}${(row)})`,
                    date1904: false,
                });
            });
            
            /* Saldo total */
            totalRowData.push({
                formula: `=SUM(${colEnd}${(rowIni + 1)}:${colEnd}${(row)})`,
                date1904: false
            });
            const totalRow = worksheet.addRow(totalRowData);

            /* Estilo de Totales */
            totalRow.eachCell((cell) => {
                const col = getExcelColumnLetter(+cell.col);
                cell.style = Object.assign(cell.style, STYLE_FOOT);
                if (cell.type != ExcelJS.ValueType.Null) {
                    cell.numFmt = CURRENCY_FORMAT;
                }

                if (col === 'A') {
                    cell.border = Object.assign({
                        left: {
                            style: 'thin',
                            color: { argb: 'FF666666' }
                        }
                    }, STYLE_FOOT.border);
                } else if (col === colEnd) {
                    cell.border = Object.assign({
                        right: {
                            style: 'thin',
                            color: { argb: 'FF666666' }
                        }
                    }, STYLE_FOOT.border);
                }
            });
                
            row += worksheet.rowCount + 2;
            cell = worksheet.getCell('A' + row);
            cell.value = (SOFT_NAME + ' v.' + VERSION);
            cell.style = STYLE_SOFT_VERSION;
            
            resolve(workbook.xlsx);
        });
    });
};

const createDateIntervalPerWeek = (date: string): Interval[] => {
    const dateStartNow = new Date(dateStringFormatSql(date));
    // const dateEndNow = new Date(dateStringFormatSql(date));

    return [
        {
            name: 'past_more_45',
            title: 'Mas de 50 días vencido',
            dateStart: null,
            dateEnd: dateStartNow.subDays(51).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) <= DATE('${dateStartNow.subDays(51).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'past_45',
            title: 'De 40 a 50 días vencido',
            dateStart: dateStartNow.subDays(50).formatSql(),
            dateEnd: dateStartNow.subDays(41).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) >= DATE('${dateStartNow.subDays(50).formatSql()}') AND DATE(cin.date_due) <= DATE('${dateStartNow.subDays(41).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'past_39',
            title: 'De 10 a 39 días vencido',
            dateStart: dateStartNow.subDays(39).formatSql(),
            dateEnd: dateStartNow.subDays(11).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) >= DATE('${dateStartNow.subDays(10).formatSql()}') AND DATE(cin.date_due) <= DATE('${dateStartNow.subDays(1).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'past_31',
            title: 'Menos de 10 días vencido',
            dateStart: dateStartNow.subDays(10).formatSql(),
            dateEnd: dateStartNow.subDays(1).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) >= DATE('${dateStartNow.subDays(10).formatSql()}') AND DATE(cin.date_due) <= DATE('${dateStartNow.subDays(1).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'past_due',
            title: 'Total vencido',
            dateStart: null,
            dateEnd: dateStartNow.subDays(1).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) <= DATE('${dateStartNow.subDays(1).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'block_1',
            title: 'Vence entre 0 a 30 días',
            dateStart: dateStartNow.formatSql(),
            dateEnd: dateStartNow.addDays(30).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) >= DATE('${dateStartNow.formatSql()}') AND DATE(cin.date_due) <= DATE('${dateStartNow.addDays(30).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'block_2',
            title: 'Semana 2',
            dateStart: dateStartNow.addDays(1).formatSql(),
            dateEnd: dateStartNow.addDays(7).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) >= DATE('${dateStartNow.addDays(1).formatSql()}') AND DATE(cin.date_due) <= DATE('${dateStartNow.addDays(7).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'block_3',
            title: 'Semana 3',
            dateStart: dateStartNow.addDays(8).formatSql(),
            dateEnd: dateStartNow.addDays(14).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) >= DATE('${dateStartNow.addDays(8).formatSql()}') AND DATE(cin.date_due) <= DATE('${dateStartNow.addDays(14).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'block_4',
            title: 'Semana 4',
            dateStart: dateStartNow.addDays(15).formatSql(),
            dateEnd: dateStartNow.addDays(21).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) >= DATE('${dateStartNow.addDays(15).formatSql()}') AND DATE(cin.date_due) <= DATE('${dateStartNow.addDays(21).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'block_5',
            title: 'Semana 5',
            dateStart: dateStartNow.addDays(22).formatSql(),
            dateEnd: dateStartNow.addDays(28).formatSql(),
            sql: `SUM(CASE WHEN DATE(cin.date_due) >= DATE('${dateStartNow.addDays(22).formatSql()}') AND DATE(cin.date_due) <= DATE('${dateStartNow.addDays(28).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
        {
            name: 'future',
            title: 'En adelante',
            dateStart: dateStartNow.addDays(29).formatSql(),
            dateEnd: null,
            sql: `SUM(CASE WHEN DATE(cin.date_due) >= DATE('${dateStartNow.addDays(29).formatSql()}') THEN (cin.amount_total/cin.currency_value) ELSE 0 END)`,
            total: 0
        },
    ];
};

export { generateAccountReceivableExcel };