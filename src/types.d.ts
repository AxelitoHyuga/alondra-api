import { Request } from "express";

export type ReportType = 'customer' | 'transaction_sequence';
export type AnalyticsRequest = Request<{}, any, any, AnalyticsFilters>;

export interface AnalyticsFilters {
    dateFrom?: string,
    dateTo: string,
    name?: string,
    origin?: string,
    customer?: string,
    salesperson?: string,
    reportType: ReportType,
    remission?: string | number
};

export interface RowAccountReceivable {
    customer_invoice_id: string,
    customer_id: string,
    customer: string,
    date_invoice: Date,
    name: string,
    origin: string,
    reference: string,
    salesperson: string,
    date_due: Date,
    total_balance: number,
    payment_term: string,
    currency_code: string
}

export interface RowAccountReceivableResponse extends RowAccountReceivable {
    past_more_45: number,
    past_45: number,
    past_31: number,
    past_due: number,
    block_1: number,
    block_2: number,
    block_3: number,
    block_4: number,
    block_5: number,
    block_6: number,
    future: number,
}

export interface Interval {
    name: string,
    title: string,
    dateStart: string | null,
    dateEnd: string | null,
    sql: string,
    total: number
}

declare global {
    interface Date {
        addDays(days: number): Date,
        subDays(days: number): Date,
        formatSql(): string,
    };
    declare var SOFT_NAME: String;
    declare var VERSION: String;
}
