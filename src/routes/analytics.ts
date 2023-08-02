import express from "express";
import 'dotenv/config';
import { AnalyticsRequest, CustomerInvoiceFilters } from "../types";
import { generateAccountReceivableExcel, generateCustomerInvoiceExcel } from "../services/analyticsService";
import { CustomError } from "../tools";
const analyticsRouter = express.Router();

analyticsRouter.get('/cuentas_por_cobrar.xlsx', async(req: AnalyticsRequest, res) => {
    const data = req.query;
    const fileName = 'cuentas_por_cobrar.xlsx';

    try {
        const xlsx = await generateAccountReceivableExcel(data);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader("Content-Disposition", "attachment; filename=" + fileName);
        await xlsx.write(res);
        res.end();
    } catch(err) {
        console.error(err);
        if (err instanceof CustomError) {
            res.status(err.code).send(err.message);
        }
    }
});

analyticsRouter.get('/reporte_facturacion_clientes.xlsx', async(req: AnalyticsRequest, res) => {
    const data = req.query;
    const fileName = 'cuentas_por_cobrar.xlsx';

    try {
        const xlsx = await generateCustomerInvoiceExcel(data as CustomerInvoiceFilters);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader("Content-Disposition", "attachment; filename=" + fileName);
        await xlsx.write(res);
        res.end();
    } catch(err) {
        console.error(err);
        if (err instanceof CustomError) {
            res.status(err.code).send(err.message);
        }
    }
});

export default analyticsRouter;