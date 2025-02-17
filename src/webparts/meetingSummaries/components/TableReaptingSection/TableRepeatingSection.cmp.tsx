import React, { useState, useEffect, memo } from 'react';
import styles from './TableReaptingSection.module.scss';
import Title from '../Title/Title.cmp';
import { Schema } from '../Interfaces';
import Fab from '@mui/material/Fab';
import AddIcon from '@mui/icons-material/Add';
import { TextField, Autocomplete, Select, MenuItem, Tooltip } from '@mui/material';
import { DatePicker } from '@mui/x-date-pickers/DatePicker';
import DeleteIcon from '@mui/icons-material/Delete';
import { IconButton } from '@mui/material';
import AddBoxIcon from '@mui/icons-material/AddBox';
import { v4 as uuidv4 } from 'uuid';
import { UnifiedNameAutocomplete } from '../UnifiedNameAutocomplete/UnifiedNameAutocomplete.cmp';
import PopUp from '../PopUp/PopUp.cmp';
import RichText from '../RichText/RichText.cmp';
import EditNoteIcon from '@mui/icons-material/EditNote';
import ReactQuill from 'react-quill';
import 'react-quill/dist/quill.snow.css'; // Import Quill.js default styles

interface TableTableRepeatingSection {
    label: string;
    name: string;
    schema: Schema;
    data: Record<string, any>[]; // Ensure each row has a unique `uid`
    addRow?: () => void;
    attachRow?: (rowIndex: number) => void;
    deleteRow?: (rowIndex: number) => void
    onChangeGeneric: any
    context: any,
    companies: any,
    users: any,
    currDir: boolean
    errors?: any
}

const noBorder = {
    "& .MuiOutlinedInput-notchedOutline": {
        border: "none",
        "&:hover": {
            border: "none",
        },
        "&.Mui-focused": {
            border: "none",
        },
    },

}

const TableRepeatingSection = memo(function TableRepeatingSection({
    schema,
    label,
    name,
    data,
    addRow,
    attachRow,
    deleteRow,
    onChangeGeneric,
    context,
    companies,
    users,
    currDir,
    errors
}: TableTableRepeatingSection) {

    const [localData, setLocalData] = useState(data);
    const [isDialogOpen, setIsDialogOpen] = useState(false);
    const [editingRowId, setEditingRowId] = useState<string | null>(null);
    const [currentRichTextValue, setCurrentRichTextValue] = useState<string>('');

    // Synchronize internal state when the parent data changes
    useEffect(() => {
        setLocalData(data);
    }, [data]);

    const handleLocalChange = (e: any, rowIndex: number, fieldName: string) => {
        const value = e.target.value;

        // Update internal state first
        setLocalData((prev) =>
            prev.map((row: any, index: number) =>
                index === rowIndex ? { ...row, [fieldName]: value } : row
            )
        );
    };

    const handleBlur = (e: any, rowIndex: number, fieldName: string) => {
        // Call parent state update after user interaction finishes
        onChangeGeneric(e, name, fieldName, rowIndex);
    };

    const handleOpenPopup = (rowId: string, initialValue: string) => {
        setEditingRowId(rowId);
        setCurrentRichTextValue(initialValue || '');
        setIsDialogOpen(true);
    };

    const handleSavePopup = (value: string) => {
        if (editingRowId) {
            setLocalData((prev) =>
                prev.map((row) =>
                    row.uid === editingRowId ? { ...row, description: value } : row
                )
            );

            onChangeGeneric(
                { target: { value } },
                name,
                'description',
                localData.findIndex((row) => row.uid === editingRowId)
            );

            setIsDialogOpen(false);
            setEditingRowId(null);
        }
    };

    const renderRichText = (htmlContent: string) => (
        <div
            dangerouslySetInnerHTML={{ __html: htmlContent }}
            style={{
                padding: '5px',
                minHeight: '40px',
            }}
        />
    );


    const renderField = (field: any, value: any, rowIndex: number, dataArrayName: string) => {
        if (field.name === 'id') {
            return <span>{`${value}.`}</span>;
        }
        if (field.name === "idView") {
            return <TextField sx={noBorder} value={value || ""} onChange={(e) => onChangeGeneric(e, dataArrayName, field.name, rowIndex)}></TextField>
        } else if (field.name === "importance" || field.name === 'status') {
            const choices: string[] = field.name === 'importance' ? ["Low", "Medium", "High", "Urgent"] : ["Not started", "In progress", "Done", "Canceled"]
            return (
                <Select
                    sx={noBorder}
                    fullWidth
                    value={value || ""} // Ensure value is controlled
                    onChange={(e) => onChangeGeneric(e, dataArrayName, field.name, rowIndex)}
                >
                    {choices.map((name, index) => (
                        <MenuItem key={index} value={name}>
                            {name}
                        </MenuItem>
                    ))}
                </Select>
            );
        } else if (field.name === "company") {
            return (
                <Autocomplete
                    sx={noBorder}
                    className={styles.fullWidth}
                    options={companies} // List of options
                    freeSolo
                    value={value || ""} // Ensure controlled value
                    onChange={(event, newValue) => {
                        // Handle option selection or freeSolo input
                        const finalValue = typeof newValue === "string" ? newValue : newValue?.toString() || "";
                        handleLocalChange(
                            { target: { value: finalValue } },
                            rowIndex,
                            field.name
                        );
                    }}
                    onBlur={(event) => {
                        // Update value on blur
                        onChangeGeneric(event, dataArrayName, field.name, rowIndex);
                    }}
                    renderInput={(params) => (
                        <Tooltip title="">
                            <TextField
                                {...params}
                                variant="outlined"
                            />
                        </Tooltip>
                    )}
                />
            )
        } else if (field.name === "startDate" || field.name === "endDate" || field.name === "dueDate") {
            return (
                <DatePicker
                    sx={noBorder}
                    className={styles.fullWidth}
                    value={value || null} // Ensure value is controlled                    
                    minDate={
                        field.name === 'endDate' && localData?.[rowIndex]?.startDate
                            ? localData[rowIndex].startDate
                            : null
                    }
                    maxDate={field.name === 'startDate' && localData?.[rowIndex]?.endDate
                        ? localData[rowIndex].endDate
                        : null
                    }
                    onChange={(newValue) =>
                        onChangeGeneric({ target: { value: newValue } }, dataArrayName, field.name, rowIndex)
                    }
                />
            );
        } else if (field.name === "name") {
            return (
                <UnifiedNameAutocomplete
                    sx={noBorder}
                    context={context}
                    params={{ id: rowIndex }}
                    users={users}
                    value={Array.isArray(value) ? value : value ? [value] : []}
                    multiple={true}
                    freeSolo={true}
                    label='Name'
                    errors={errors}
                    row={rowIndex}
                    currDir={currDir}
                    onChange={(id: any, newValue: any, email: any, isFreeSolo: boolean, ids: string[], onBlur: string) => {

                        onChangeGeneric(
                            { target: { value: newValue, Email: email, isFreeSolo, ids } },
                            dataArrayName,
                            field.name,
                            rowIndex,
                            onBlur
                        );
                    }}
                />
            );

        } else if (field.name === "forInfo") {
            const filterdUsers = users.filter((u: any) => u?.Email)
            return (
                <UnifiedNameAutocomplete
                    sx={noBorder}
                    context={context}
                    params={{ id: rowIndex }}
                    users={filterdUsers}
                    value={Array.isArray(value) ? value : value ? [value] : []}
                    multiple={true}
                    freeSolo={true}
                    label='ForInfo'
                    onChange={(id: any, newValue: any, email: any, isFreeSolo: boolean, forInfoIds: string[], onBlur: string) => {
                        onChangeGeneric(
                            { target: { value: newValue, Email: email, isFreeSolo, forInfoIds } },
                            dataArrayName,
                            field.name,
                            rowIndex,
                            onBlur
                        );
                    }}
                />
            );
        } else if (field.name === "action") {
            return (
                <IconButton size="small" sx={{ display: "flex", justifyContent: "center" }} onClick={() => deleteRow?.(rowIndex)}>
                    <DeleteIcon />
                </IconButton>
            );
        } else if (field.name === "attach") {
            return (
                <IconButton size="small" sx={{ display: "flex", justifyContent: "center" }} onClick={() => attachRow?.(rowIndex)}>
                    <AddBoxIcon />
                </IconButton>
            );
        } else if (field.name === "description") {
            return (
                // <div style={{ display: 'flex', width: '100%' }}>
                //     <div style={{ flexGrow: 1, marginRight: '8px' }}>
                //         {renderRichText(value || '')}
                //     </div>
                //     <div >

                //         <IconButton
                //             onClick={() => handleOpenPopup(localData[rowIndex].uid, value)}
                //             style={{ color: 'gray' }}
                //         >
                //             <EditNoteIcon fontSize="small" />
                //         </IconButton>
                //     </div>
                // </div>
                <div style={{ display: 'flex', width: '100%' }}>
                    <div style={{ flexGrow: 1, marginRight: '8px' }}>
                        <TextField
                            sx={noBorder}
                            multiline
                            fullWidth
                            value={value || ""} // Render plain text or a placeholder
                            onChange={(e) => handleLocalChange(e, rowIndex, "description")} // Update local state
                            onBlur={(e) => handleBlur(e, rowIndex, "description")} // Trigger parent state update
                            style={{
                                padding: '5px',
                                minHeight: '40px',
                            }}
                        />
                    </div>
                    <div>
                        <IconButton
                            onClick={() => handleOpenPopup(localData[rowIndex].uid, value)} // Open popup for detailed editing
                            style={{ color: 'gray' }}
                        >
                            <EditNoteIcon fontSize="small" />
                        </IconButton>
                    </div>
                </div>

            );
        }
        return (
            <TextField
                sx={noBorder}
                fullWidth
                value={value || ""}
                onChange={(e) => handleLocalChange(e, rowIndex, field.name)}
                onBlur={(e) => handleBlur(e, rowIndex, field.name)}
            />
        );
    };

    return (
        <div className={styles.tableRepeatingSection}>
            <Title label={label} />

            {/* Popup for Rich Text Editing */}
            {/* <PopUp open={isDialogOpen} onClose={() => setIsDialogOpen(false)}>
                <ReactQuill
                    value={currentRichTextValue}
                    onChange={setCurrentRichTextValue}
                    theme="snow"
                />
                <button
                    onClick={() => handleSavePopup(currentRichTextValue)}
                    style={{
                        marginTop: '10px',
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
            </PopUp> */}

            <PopUp open={isDialogOpen} onClose={() => setIsDialogOpen(false)} dir={!currDir ? 'ltr' : 'rtl'}>
                <textarea
                    value={currentRichTextValue}
                    onChange={(e) => setCurrentRichTextValue(e.target.value)} // Update text value
                    style={{
                        width: '100%',
                        height: '150px',
                        padding: '10px',
                        fontSize: '14px',
                        border: '1px solid #ccc',
                        borderRadius: '4px',
                        resize: 'vertical',
                        direction: !currDir ? 'ltr' : 'rtl'
                    }}
                    placeholder={!currDir ? "Enter description..." : 'הכנס תיאור'}
                />
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center' }}>
                    <button
                        onClick={() => handleSavePopup(currentRichTextValue)}
                        style={{
                            marginTop: '10px',
                            padding: '5px 10px',
                            fontSize: '14px',
                            backgroundColor: '#1976d2',
                            color: '#fff',
                            border: 'none',
                            borderRadius: '4px',
                            cursor: 'pointer',
                            direction: !currDir ? 'ltr' : 'rtl'
                        }}
                    >
                        {!currDir ? 'Save' : 'שמור'}
                    </button>
                </div>
            </PopUp>


            <div className={styles.tableWrapper}>
                <div className={`${name === 'tasks' ? styles.largeTable : name === 'meetingContent' ? styles.meetingContentTable : styles.table}`}>
                    {/* Table Header */}
                    <div className={styles.header}>
                        {schema.fields.map((field) => (
                            <div
                                key={field.name || uuidv4()} // Stable and unique key
                                className={styles.headerCell}
                                style={{
                                    maxWidth: field.width || 'none',
                                    minWidth: field.width || 200,
                                }}
                            >
                                {field.label}
                            </div>
                        ))}
                    </div>
                    {/* Table Rows */}
                    {localData.map((rowData, rowIndex) => (
                        <div key={rowData.uid} className={styles.row}>
                            {schema.fields.map((field) => (
                                <div
                                    key={`${rowData.uid}-${field.name}`} // Stable and unique key
                                    className={styles.cell}
                                    style={{
                                        maxWidth: field.width || 'none',
                                        minWidth: field.width || 200,
                                    }}
                                >
                                    {renderField
                                        ? renderField(field, rowData[field.name], rowIndex, name)
                                        : rowData[field.name] || '-'} {/* Default rendering */}
                                </div>
                            ))}
                        </div>
                    ))}
                </div>
            </div>

            <div style={{ display: 'flex', justifyContent: 'center', padding: '1em' }}>
                <Fab size="small" aria-label="add" color='success' sx={{ backgroundColor: '#8AC693' }} onClick={addRow}>
                    <AddIcon htmlColor="white" />
                </Fab>
            </div>
        </div >
    );
})

export default TableRepeatingSection;
