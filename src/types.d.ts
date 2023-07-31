import { Request } from "express";

export type ReportType = 'customer' | 'transaction_sequence' | 'detail';
export type AnalyticsRequest = Request<{}, any, any, AnalyticsFilters | CustomerInvoiceFilters>;

export interface AnalyticsFilters {
    dateFrom?: string,
    dateTo: string,
    name?: string,
    origin?: string,
    customer?: string | number | Array,
    salesperson?: string | number,
    reportType: ReportType,
    remission?: string | number,
};

export interface CustomerInvoiceFilters extends AnalyticsFilters {
    productSearch?: string | Array,
    productCategoryId?: string | number,
    productManufaturerId?: string | number,
    reference?: string,
    transactionSequenceId?: string | number,
    invoiceStatusId?: string | number,
    invoice?: boolean | number,
    promotionId?: string | number,
    showCanceled?: string | number,
    productRange?: string,
    customerRange?: string,
    sort?: string,
    order?: string,
    marginPermission: string | number,
    configCurrencyCode: string,
}

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

export interface RowCustomerInvoice {
    total_count: number,
    total_quantity: number,
    total_amount_untaxed: number,
    total_amount_untaxed_withoutdiscount: number,
    total_amount_discount: number,
    total_amount_tax: number,
    total_amount_tax_ret: number,
    total_amount_tax_total: number,
    total_amount_total: number,
    total_amount_cost: number,
    total_amount_margin: number,
    total_percent_margin: number,
    date_invoice: Date,
    customer_invoice_id: string,
    customer_id: string,
    number: string,
    manufacturer: string,
    name: string,
    date_due: Date,
    origin: string,
    reference: string,
    currency_code: string,
    invoice_status_id: string,
    remission: string,
    salesperson: string,
    payment_term: string,
    invoice_status: string,
    transaction_sequence: string,
    uuid: string,
    default_code: string,
    product: string,
    product_uom_id: string,
    decimal_place_qty: string,
    product_category: string,
    customer: string,
    customer_invoice_line_id: string,
    comission_rate: string,
    comission_rate_product: string,
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
