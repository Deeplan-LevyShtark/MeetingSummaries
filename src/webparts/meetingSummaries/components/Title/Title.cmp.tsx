import * as React from 'react';
import styles from './Title.module.scss';

interface TitleProps {
    label: string
}

export default function Title({ label }: TitleProps) {

    return (
        <h3 className={styles.sectionHeader}>{label}</h3>
    );
}