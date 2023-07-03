/**
 * 
 * @param date Fecha
 * @returns Fecha formateada
 */
export function dateStringFormatSql(date: string): string {
    const reverse = date.split('-').reverse();
    const nDate = new Date(reverse.join('/'));
    const year = new Intl.DateTimeFormat('en', { year: 'numeric' }).format(nDate);
    const month = new Intl.DateTimeFormat('en', { month: '2-digit' }).format(nDate);
    const day = new Intl.DateTimeFormat('en', { day: '2-digit' }).format(nDate);

    return `${year}-${month}-${day}`;
}

export function configure() {
    /* Añadir días a una fecha */
    Date.prototype.addDays = function(days: number): Date {
        const date = new Date(this.valueOf());
        date.setDate(date.getDate() + days);
        return date;
    }

    /* Eliminar días a una fecha */
    Date.prototype.subDays = function(days: number): Date {
        const date = new Date(this.valueOf());
        date.setDate(date.getDate() - days);
        return date;
    }

    /* Formatea la fecha */
    Date.prototype.formatSql = function(): string {
        const date = new Date(this.valueOf());
        const year = new Intl.DateTimeFormat('en', { year: 'numeric' }).format(date);
        const month = new Intl.DateTimeFormat('en', { month: '2-digit' }).format(date);
        const day = new Intl.DateTimeFormat('en', { day: '2-digit' }).format(date);

        return `${year}-${month}-${day}`;
    }

    /* Variables globales */
    globalThis.SOFT_NAME = 'ALA';
    globalThis.VERSION = '7.3.1'
}

export function getExcelColumnLetter(columnNumber: number): string {
    let dividend = columnNumber;
    let columnName = '';

    while (dividend > 0) {
        const modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(65 + modulo) + columnName;
        dividend = Math.floor((dividend - modulo) / 26);
    }

    return columnName;
}