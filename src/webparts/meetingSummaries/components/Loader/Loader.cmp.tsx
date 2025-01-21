import * as React from 'react';
import styles from './Loader.module.scss';

interface LoaderProps {

}

export default function Loader({ }: LoaderProps) {

    return (
        <div className={styles.loaderContainer}>

            <div className={styles.loadingspinner}>
                <div id={styles.square1}></div>
                <div id={styles.square2}></div>
                <div id={styles.square3}></div>
                <div id={styles.square4}></div>
                <div id={styles.square5}></div>
            </div>

        </div>
    );
}