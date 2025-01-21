import React, { useState } from 'react';
import ReactQuill from 'react-quill'; // Import Quill.js
import 'react-quill/dist/quill.snow.css'; // Quill.js default styling
import { createTheme, ThemeProvider } from '@mui/material/styles';
import styles from './RichText.module.scss'; // Keep your existing styles

const theme = createTheme({
    typography: {
        fontFamily: 'Roboto, Arial, sans-serif',
    },
});

const RichText: React.FC<{ initialValue: string; onSave: (value: string) => void }> = ({ initialValue, onSave }) => {
    const [value, setValue] = useState(initialValue || ''); // Store editor content

    // Save content on button click
    const handleSave = () => {
        onSave(value); // Send the HTML content to the parent
    };

    return (
        <ThemeProvider theme={theme}>
            <div className={styles.richTextContainer}>
                {/* Editor */}
                <div
                    style={{
                        border: '1px solid #ccc',
                        borderRadius: '4px',
                        padding: '10px',
                        marginBottom: '10px',
                    }}
                >
                    <ReactQuill
                        value={value}
                        onChange={setValue}
                        theme="snow" // Quill's default theme
                        placeholder="Type here..."
                        modules={{
                            toolbar: [
                                ['bold', 'italic', 'underline'], // Formatting buttons
                                [{ list: 'ordered' }, { list: 'bullet' }], // Lists
                                ['clean'], // Remove formatting
                            ],
                        }}
                    />
                </div>

                {/* Save Button */}
                <button
                    onClick={handleSave}
                    style={{
                        padding: '5px 10px',
                        fontSize: '14px',
                        backgroundColor: '#1976d2',
                        color: '#fff',
                        border: 'none',
                        borderRadius: '4px',
                        cursor: 'pointer',
                    }}
                >
                    Save
                </button>
            </div>
        </ThemeProvider>
    );
};

export default RichText;
