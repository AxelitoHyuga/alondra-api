import express from "express";
import 'dotenv/config';
import { AnalyticsRequest } from "../types";
import { generateAccountReceivableExcel } from "../services/analyticsService";
const analyticsRouter = express.Router();

analyticsRouter.get('/cuentas_por_cobrar.xlsx', (req: AnalyticsRequest, res) => {
    const data = req.query;
    const fileName = 'cuentas_por_cobrar.xlsx';

    try {
        generateAccountReceivableExcel(data).then((xlsx) => {
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader("Content-Disposition", "attachment; filename=" + fileName);
            xlsx.write(res).then(() => {
                res.end();
            });
        });
    } catch(err) {
        console.error(err);
        res.status(500).send();
    }
});

export default analyticsRouter;