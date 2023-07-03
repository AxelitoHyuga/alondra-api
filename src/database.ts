import { createConnection } from "mysql";

const connection = createConnection({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_DATABASE,
    typeCast: (field, next) => {
        if (field.type === 'DATE') {
            const value: string | null = field.string();
            if (value === null) {
                return value;
            }
            return new Date(value);
        }

        return next();
    }
});

connection.connect((err) => {
    if (err) {
        console.log(err.stack);
        return;
    }
    console.log('Database connected!');
});

export default connection;