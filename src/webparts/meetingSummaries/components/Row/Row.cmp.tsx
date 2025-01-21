import * as React from 'react';
import styles from './Row.module.scss';
import { Schema } from '../Interfaces';
import Cell from '../Cell/Cell.cmp';
import { TextField } from '@mui/material';

interface RowProps {
    data: Record<string, any>;
    schema: Schema;
}

export default function Row({ data, schema }: RowProps) {

    return (
        <div className={styles.rowContainer}>
            {schema.fields.map((field) => (
                <Cell label="Name" editable={true}>
                    <TextField fullWidth ></TextField>
                </Cell>
            ))}
        </div>
    );
}