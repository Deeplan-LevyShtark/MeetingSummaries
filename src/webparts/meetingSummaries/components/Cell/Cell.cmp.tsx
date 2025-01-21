import * as React from 'react';
import styles from './Cell.module.scss';

interface CellProps {
    label: string;
    editable?: boolean;
    children?: React.ReactNode; // Allows passing any JSX as children
}

export default function Cell({ label, editable, children }: CellProps) {
    return (
        <div className={styles.cell}>
            {editable ? (
                children // Render the passed component as children
            ) : (
                <span>{label}</span>
            )}
        </div>
    );
}
