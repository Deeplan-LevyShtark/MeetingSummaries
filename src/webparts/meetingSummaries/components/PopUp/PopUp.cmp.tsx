import * as React from 'react';
import styles from './PopUp.module.scss';
import { Dialog, DialogTitle, DialogContent, DialogActions, Button } from '@mui/material';

interface PopUpProps {
    open: boolean; // Controls whether the dialog is open or closed
    onClose: () => void; // Callback to handle closing the dialog
    title?: string; // Optional title for the dialog
    children: React.ReactNode; // Generic children to render inside the dialog
    actions?: React.ReactNode; // Optional actions for the dialog footer
    dir?: string
}

export default function PopUp({ open, onClose, title, children, actions, dir }: PopUpProps) {
    return (
        <div className={styles.popUpContainer}>
            <Dialog open={open} onClose={onClose} maxWidth="md" fullWidth dir={dir}>
                {title && <DialogTitle>{title}</DialogTitle>}
                <DialogContent>{children}</DialogContent>
                {actions && <DialogActions>{actions}</DialogActions>}
            </Dialog>
        </div>
    );
}
