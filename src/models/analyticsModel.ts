import connection from "../database";
import { dateStringFormatSql } from "../tools";
import { AnalyticsFilters, CustomerInvoiceFilters, Interval, RowAccountReceivable } from "../types";
const PREFIX = process.env.DB_PREFIX;

const getCustomerAccountReceivables = (filters: AnalyticsFilters, intervals: Interval[]): Promise<RowAccountReceivable[]> => {
    return new Promise((resolve, reject) => {
        let sql = `SELECT SQL_CALC_FOUND_ROWS
        COUNT(cin.customer_invoice_id) AS total_count,
        SUM(cin.amount_untaxed/cin.currency_value*trs.nature) as total_amount_untaxed,
        SUM(cin.amount_tax/cin.currency_value*trs.nature) as total_amount_tax,
        SUM(cin.amount_tax_ret/cin.currency_value*trs.nature) as total_amount_tax_ret,
        SUM( (cin.amount_tax/cin.currency_value*trs.nature) + (cin.amount_tax_ret/cin.currency_value*trs.nature) ) as total_amount_tax_total,
        SUM(cin.amount_total/cin.currency_value*trs.nature) as total_amount_total,
        SUM(0) AS total_amount_cost,
        SUM(cin.amount_margin/cin.currency_value*trs.nature) as total_amount_margin,
        SUM(cin.amount_margin/cin.amount_untaxed/cin.currency_value * 100*trs.nature) as total_percent_margin,
        SUM(cin.balance/cin.currency_value*trs.nature) as total_balance,
        cin.date_invoice, GROUP_CONCAT(cin.customer_invoice_id) AS customer_invoice_id,
        cin.customer_id,cin.number,cin.\`name\`,cin.date_due,cin.origin,cin.reference,cin.currency_code,cin.invoice_status_id,cin.remission,
        sap.name AS salesperson,
        pte.\`name\` AS payment_term,
        ist.\`name\` AS invoice_status,trs.\`name\` AS transaction_sequence,
        cfd.uuid,
        '' as default_code,'' as product,0 as total_quantity,0 as decimal_place_qty,'' as product_category,`;

        intervals.forEach((interval) => {
            sql += `${interval.sql} AS ${interval.name}, `;
        });

        sql += `cus.name AS customer FROM ${PREFIX}customer_invoice AS cin
        INNER JOIN ${PREFIX}customer AS cus ON cin.customer_id = cus.customer_id
        INNER JOIN ${PREFIX}user AS sap ON cin.salesperson_id = sap.id
        INNER JOIN ${PREFIX}invoice_status AS ist ON cin.invoice_status_id=ist.invoice_status_id
        INNER JOIN ${PREFIX}currency as curr ON cin.currency_id=curr.currency_id
        LEFT JOIN ${PREFIX}customer_invoice_cfdi AS cfd ON cin.customer_invoice_id=cfd.customer_invoice_id
        INNER JOIN ${PREFIX}payment_term AS pte ON cin.payment_term_id=pte.payment_term_id
        INNER JOIN ${PREFIX}transaction_sequence AS trs ON cin.transaction_sequence_id=trs.id
        WHERE cin.invoice_status_id IN (2,4,5) 
        AND trs.code IN ('debit.invoice', 'dedit.fee', 'debit.lease', 'debit.debit', 'debit.remission')`;

        if (filters) {
            if (filters.dateFrom) {
                const date: string = dateStringFormatSql(filters.dateFrom);
                sql += ` AND DATE(cin.date_invoice) >= DATE('${ date }')`;
            }
            if (filters.dateTo) {
                const date: string = dateStringFormatSql(filters.dateTo);
                sql += ` AND DATE(cin.date_invoice) <= DATE('${ date }')`;
            }
            if (filters.name) {
                sql += ` AND cin.name LIKE '%${ filters.name.replace(/\s/gm, '%%') }%'`;
            }
            if (filters.origin) {
                sql += ` AND cin.origin LIKE '%${ filters.origin.replace(/\s/gm, '%%') }%'`;
            }
            if (filters.customer && typeof filters.customer == 'string') {
                sql += ` AND cus.name LIKE '%${ filters.customer.replace(/\s/gm, '%%') }%'`;
            }
            if (filters.salesperson && typeof filters.salesperson == 'string') {
                sql += ` AND sap.name LIKE '%${ filters.salesperson.replace(/\s/gm, '%%') }%'`;
            }
            if (filters.remission) {
                sql += ` AND cin.remission = ${ filters.remission }`;
            }
        }

        sql += ` GROUP BY cin.customer_invoice_id ORDER BY cus.\`name\`, cin.date_due ASC`;

        /* Realiza la consulta */
        connection.query({ sql }, (err, rows, _fields) => {
            if (err)
                reject(err);

            resolve(rows);
        });
    });
};

const getCustomerPaymentReconcilied = (customerId: string): Promise<any[]> => {
    return new Promise((resolve, reject) => {
        let sql = `(SELECT cpr.*, cpa.\`name\` as customer_credit, DATE_FORMAT(cpa.date_effective,'%Y-%m-%d') as customer_credit_date_credit,cpa.amount as customer_credit_amount_total,
            cpa.currency_code as customer_credit_currency_code, cin.\`name\` AS customer_invoice,cin.date_due as customer_invoice_date_due,cin.amount_total as customer_invoice_amount_total,
            cin.currency_code as customer_invoice_currency_code, cpr.date_conciled as customer_date_conciled
        FROM ${PREFIX}customer_payment_invoice_rel AS cpr
            INNER JOIN ${PREFIX}customer_payment AS cpa ON cpr.customer_payment_id=cpa.customer_payment_id
            INNER JOIN ${PREFIX}customer_invoice AS cin ON cpr.customer_invoice_id=cin.customer_invoice_id
        WHERE cin.customer_id = ${customerId} AND cpr.status = 1)
        UNION ALL
        (SELECT crr.*,
            crf.\`name\` as customer_credit,DATE_FORMAT(crf.date_invoice,'%Y-%m-%d') as customer_credit_date_credit,crf.amount_total as customer_credit_amount_total,
            crf.currency_code as customer_credit_currency_code, cin.\`name\` AS customer_invoice,cin.date_due as customer_invoice_date_due,cin.amount_total as customer_invoice_amount_total,
            cin.currency_code as customer_invoice_currency_code, crr.date_conciled as customer_date_conciled
        FROM ${PREFIX}customer_refund_invoice_rel AS crr
            INNER JOIN ${PREFIX}customer_invoice AS crf ON crr.customer_refund_id=crf.customer_invoice_id
            INNER JOIN ${PREFIX}customer_invoice AS cin ON crr.customer_invoice_id=cin.customer_invoice_id
        WHERE cin.customer_id = ${customerId} AND crr.status = 1)
        ORDER BY customer_date_conciled`;

        connection.query(sql, (err, rows) => {
            if (err)
                reject(err);

            resolve(rows);
        });
    });
};

const getCustomerInvoiceLinesReport = (filters: CustomerInvoiceFilters, group?: number, sort?: number, order?: string): Promise<any[]> => {
    return new Promise((resolve, reject) => {
        const groupBy = [
            'DATE(cin.date_invoice)',
            'MONTH(cin.date_invoice)',
            'cin.customer_id',
            'cin.customer_invoice_id',
            'cli.customer_invoice_line_id',
        ];
        const sortBy = [
            'DATE(cin.date_invoice)',
            'cus.`name`',
        ];

        let sql = `SELECT SQL_CALC_FOUND_ROWS
        COUNT(cli.customer_invoice_line_id) AS total_count,
        SUM(cli.quantity) AS total_quantity,
        SUM(cli.price_subtotal/cin.currency_value*trs.nature) AS total_amount_untaxed,
        SUM((cli.price_subtotal/cin.currency_value*trs.nature)/((100 - cli.discount)/100)) AS total_amount_untaxed_withoutdiscount,
        SUM((cli.price_subtotal/cin.currency_value*trs.nature)*(cli.discount/100)/((100 - cli.discount)/100)) AS total_amount_discount,
        SUM(cli.price_tax/cin.currency_value*trs.nature) AS total_amount_tax,
        SUM(cli.price_tax_ret/cin.currency_value*trs.nature) AS total_amount_tax_ret,
        SUM((cli.price_tax/cin.currency_value*trs.nature) + (cli.price_tax_ret/cin.currency_value*trs.nature)) AS total_amount_tax_total,
        SUM(cli.price_total/cin.currency_value*trs.nature) AS total_amount_total,
        SUM(cli.cost/cin.currency_value*trs.nature) AS total_amount_cost,
        SUM(cli.margin/cin.currency_value*trs.nature) AS total_amount_margin,
        SUM(cli.margin/cli.price_subtotal/cin.currency_value * 100*trs.nature) AS total_percent_margin,

        cin.date_invoice,cin.customer_invoice_id,cin.customer_id,cin.number,manu.name AS manufacturer,cin.\`name\`,cin.date_due,cin.origin,cin.reference,cin.currency_code,cin.invoice_status_id,cin.remission,
        sap.name AS salesperson,
        pte.\`name\` AS payment_term,
        ist.\`name\` AS invoice_status,trs.\`name\` AS transaction_sequence,
        cfd.uuid,
        pro.default_code,cli.\`name\` AS product,pro.product_uom_id,uom.decimal_place AS decimal_place_qty,prc.\`name\` AS product_category,
        cus.\`name\` AS customer,
        cli.customer_invoice_line_id,
        sap.comission_rate,
        pro.comission_rate as comission_rate_product
        FROM ${PREFIX}customer_invoice_line AS cli
            LEFT JOIN ${PREFIX}product_uom AS uom ON cli.product_uom_id=uom.product_uom_id
            LEFT JOIN ${PREFIX}product AS pro ON cli.product_id=pro.product_id
            LEFT JOIN ${PREFIX}product_category AS prc ON pro.product_category_id=prc.product_category_id 
            LEFT JOIN ${PREFIX}manufacturer AS manu ON pro.manufacturer_id=manu.manufacturer_id 
            LEFT JOIN ${PREFIX}stock_production_lot AS spl ON cli.production_lot_id=spl.production_lot_id
            LEFT JOIN ${PREFIX}order_line AS oli ON cli.order_line_id=oli.order_line_id
            INNER JOIN ${PREFIX}customer_invoice AS cin ON cli.customer_invoice_id= cin.customer_invoice_id
            INNER JOIN ${PREFIX}location AS loc ON cin.location_id=loc.location_id
            INNER JOIN ${PREFIX}customer AS cus ON cin.customer_id=cus.customer_id
            INNER JOIN ${PREFIX}user AS sap ON cin.salesperson_id=sap.id
            INNER JOIN ${PREFIX}invoice_status AS ist ON cin.invoice_status_id=ist.invoice_status_id
            INNER JOIN ${PREFIX}currency AS curr ON cin.currency_id=curr.currency_id
            LEFT JOIN ${PREFIX}customer_invoice_cfdi AS cfd ON cin.customer_invoice_id=cfd.customer_invoice_id
            INNER JOIN ${PREFIX}payment_term AS pte ON cin.payment_term_id=pte.payment_term_id
            INNER JOIN ${PREFIX}transaction_sequence AS trs ON cin.transaction_sequence_id=trs.id
        WHERE cin.customer_invoice_id > 0`;

        /* Filtros */
        if (filters) {
            if (filters.name) {
                sql += ` AND cin.name LIKE '%${ filters.name.replace(' ', '%%') }'`;
            }
            if (filters.dateFrom) {
                const date: string = dateStringFormatSql(filters.dateFrom);
                sql += ` AND DATE(cin.date_invoice) >= DATE('${ date }')`;
            }
            if (filters.dateTo) {
                const date: string = dateStringFormatSql(filters.dateTo);
                sql += ` AND DATE(cin.date_invoice) <= DATE('${ date }')`;
            }
            if (filters.customer) {
                if (Array.isArray(filters.customer)) {
                    sql += ` AND cin.customer_id >= ${filters.customer[0]} AND cin.customer_id <= ${filters.customer[1]}`;
                } else {
                    sql += ` AND cus.name LIKE '%${ filters.customer.replace('', '%%') }%'`;
                }
            }
            if (filters.reference) {
                sql += ` AND cin.reference LIKE '%${ filters.reference.replace(' ', '%%') }%'`;
            }
            if (filters.salesperson) {
                sql += ` AND cin.salesperson_id=${ filters.salesperson }`;
            }
            if (filters.origin) {
                sql += ` AND cin.origin LIKE '%${ filters.origin.replace(' ', '%%') }'`;
            }
            if (filters.invoice) {
                sql += " AND cin.invoice='1' ";
            }
            if (filters.remission) {
                sql += ` AND cin.remission='${ filters.remission }'`;
            }
            if (filters.promotionId) {
                sql += ` AND cli.promotion_id IN (${ filters.promotionId })`;
            }
            if (filters.transactionSequenceId) {
                sql += ` AND cin.transaction_sequence_id IN (${ filters.transactionSequenceId })`;
            }
            if (filters.invoiceStatusId) {
                sql += ` AND cin.invoice_status_id IN (${ filters.invoiceStatusId })`;
            }
            if (filters.productSearch) {
                if (Array.isArray(filters.productSearch)) {
                    sql += ` AND pro.product_id >= ${filters.productSearch[0]} AND pro.product_id <= ${filters.productSearch[1]}`;
                } else {
                    const search = filters.productSearch.replace(/\s/gm, '%%');
                    sql += ` AND ( cli.\`name\` LIKE '%${ search }%'
                    OR pro.\`name\` LIKE '%${ search }%'
                    OR pro.description LIKE '%${ search }%'
                    OR pro.default_code LIKE '%${ search }%')`;
                }
            }
            if (filters.productCategoryId) {
                sql += ` AND (pro.product_category_id = ${filters.productCategoryId} OR prc.parent_id = ${filters.productCategoryId})`;
            }
            if (filters.productManufacturerId) {
                sql += ` AND pro.manufacturer_id = ${filters.productManufacturerId}`;
            }
        }

        if (group) {
            sql += ` ${groupBy[group] ? `GROUP BY ${groupBy[group]}` : ''}`;
        }

        if (sort) {
            sql += ` ORDER BY ${sortBy[sort] ? sortBy[sort] : 'cin.date_invoice,cin.\`name\`'}`;
            
            if (order) {
                sql += ` ${order}`;
            }
        }

        connection.query(sql, (err, rows) => {
            if (err)
                reject(err);

            resolve(rows);
        });
    });
}

export { getCustomerAccountReceivables, getCustomerPaymentReconcilied, getCustomerInvoiceLinesReport };