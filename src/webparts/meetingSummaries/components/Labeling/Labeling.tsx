import React, { useEffect, useState } from 'react';
import { Autocomplete, Button, Fab, IconButton, TextField } from '@mui/material';
import { SPFI } from '@pnp/sp';
import styles from './Labeling.module.scss';
import { v4 as uuidv4 } from 'uuid';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { UnifiedNameAutocomplete } from '../UnifiedNameAutocomplete/UnifiedNameAutocomplete.cmp';
import AddIcon from '@mui/icons-material/Add';
import DeleteIcon from '@mui/icons-material/Delete';

const WP_PHASE_ARRAY = ['Infra 2', 'Alignment', '(3rd Party)', 'cordination (danon)'];
const HAS_NO_ELEMETNS = ['Infra 2', 'General']
const HAS_NO_SUB_DICIPLINES = ['General', 'Alignment']

interface LookupField {
    Id?: number | null;
    Title?: string
}

interface InputData {
    documentLibraryNameMapped: string;
    Rev?: number | null;
    RevisionNo?: number | null;
    WP?: LookupField;
    Phase?: LookupField;
    "Sub Disciplines"?: LookupField;
    Elements?: LookupField;
    "Design Stage"?: LookupField;
    "Document Status"?: LookupField;
    AuthorDesingerName?: string | null;
    Authority?: string | null
    _designType?: string | null
}

export interface LabelingProps {
    sp: SPFI;
    context: WebPartContext;
    dir: boolean;
    users: any[];
    selectedLabeling?: any; // could be an object or an array of objects
    onSave: (selectedLabeling: any[]) => void;
    onClose: () => void;
}

const mapWP: { [key: string]: string } = {
    "Wp1": "Wp1new",
    "Wp2.1": "Wp21new",
    "Wp2.2": "Wp22new",
    "Wp3": "Wp3new",
    "Wp4": "Wp4new",
    "Wp5": "Wp5new",
    "Wp6": "Wp6new",
    "Wp7": "Wp7new",
    "Wp8": "Wp8new",
    "Wp9": "Wp9new",
    "Infra 2": "Infra2new",
    "Alignment": "AlignmentNew",
    "General": "GeneralNew",
    "(3rd Party)": "3rdPartyNew",
    "cordination (danon)": "cordinationDanonNew",
};

export function Labeling(props: LabelingProps) {
    // Data from SharePoint for all lookup/autocomplete fields
    const [designData, setDesignData] = useState({
        Design_WP: [] as any[],
        Design_DesignStage: [] as any[],
        Elements: [] as any[],
        DesignDisciplinesSubDisciplines: [] as any[],
        Design_DocumentStatus: [] as any[],
        Design_Type: [] as any[],
        Phase: [] as any[],
    });

    // ROW STATE: Each row (i.e. labeling path) holds these fields.
    // Now each field will store a single object (or null), not an array.
    const [labelingArr, setLabelingArr] = useState<any[]>(() => {

        if (props.selectedLabeling) {

            return Array.isArray(props.selectedLabeling)
                ? props.selectedLabeling
                : [props.selectedLabeling];
        }
        return [
            {
                id: 1,
                WP: null,
                Phase: null,
                "Design Stage": null,
                Elements: null,
                "Sub Disciplines": null,
            },
        ];
    });

    // COMMON (STATIC) DATA across all rows
    const [commonData, setCommonData] = useState<any>(() => {
        if (props.selectedLabeling) {
            const first = Array.isArray(props.selectedLabeling)
                ? props.selectedLabeling[0]
                : props.selectedLabeling;
            return {
                Rev: first.Rev ?? 0,
                "Document Status": first["Document Status"] ?? null,
                AuthorDesingerName: first.AuthorDesingerName ?? null,
                RevisionNo: first.RevisionNo ?? 0,
                Authority: first.Authority ?? null
            };
        }
        return {
            Rev: 0,
            "Document Status": null,
            AuthorDesingerName: null,
            RevisionNo: 0,
            Authority: null
        };
    });

    // Fetch lookup data on mount
    useEffect(() => {
        getLabelingData();
    }, []);

    const getLabelingData = async () => {
        const SIZE = 5000
        try {
            const [wp, designStage, elements, disciplines, designDocumentStatus, designType, designPhase] = await Promise.all([
                props.sp.web.lists.getByTitle('Design_WP').items.select('Title, Id').top(SIZE)(),
                props.sp.web.lists.getByTitle('Design_DesignStage').items.select('Title, Id, WP_Type, Phase').top(SIZE)(),
                props.sp.web.lists.getByTitle('Elements').items.select('Title, Id, WP, Location, elementCode, elementName, elementType, ElementNameAndCode').top(SIZE)(),
                props.sp.web.lists.getByTitle('DesignDisciplinesSubDisciplines').items.select('Title, Id, Discipline, DisciplineValue, SubDiscipline').top(SIZE)(),
                props.sp.web.lists.getByTitle('Design_DocumentStatus').items.select('Title, Id').top(SIZE)(),
                props.sp.web.lists.getByTitle('Design_TYPE').items.select('Title, Id').top(SIZE)(),
                props.sp.web.lists.getByTitle('Design_Phase').items.select('Title, Id, WP_Type').top(SIZE)()
            ]);

            const filterDuplicats = (arr: any[], name: string = 'Title') => arr.filter((item, index, self) => index === self.findIndex((t) => t[name] === item[name]))

            setDesignData({
                Design_WP: filterDuplicats(wp).filter(wp => !['(3rd Party)', 'cordination (danon)'].includes(wp.Title)),
                Design_DesignStage: filterDuplicats(designStage),
                Elements: filterDuplicats(elements, 'ElementNameAndCode'),
                DesignDisciplinesSubDisciplines: filterDuplicats(disciplines, 'SubDiscipline'),
                Design_DocumentStatus: filterDuplicats(designDocumentStatus),
                Design_Type: filterDuplicats(designType).map(item => item.Title),
                Phase: filterDuplicats(designPhase),
            });
        } catch (error) {
            console.error('Error fetching labeling data:', error);
        }
    };

    // Add a new empty row to labelingArr.
    function addRow() {
        setLabelingArr((prevState: any[]) => [
            ...prevState,
            {
                id: prevState.length + 1,
                WP: null,
                Phase: null,
                "Design Stage": null,
                Elements: null,
                "Sub Disciplines": null,
            },
        ]);
    }

    function deleteRow(rowIndex: any): void {
        if (labelingArr.length === 1) return

        const filteredRows = labelingArr.filter(row => rowIndex !== row.id)

        const reformattedRows = filteredRows.map((row, index) => ({
            ...row,
            id: index + 1
        }))

        setLabelingArr(reformattedRows)
    }

    // ROW-LEVEL AUTOCOMPLETE: each row’s value comes from labelingArr[rowIndex].
    // Here, multiple selection is disabled by setting multiple to false.
    function AutoCompleteLabeling(
        rowIndex: number,
        options: any[],
        label: string, // e.g., 'WP', 'Phase', etc.
        valueField: string,
        required?: boolean,
        multiple: boolean = false
    ) {
        return (
            <div style={{ width: '100%' }}>
                <span style={{ fontFamily: 'sans-serif', paddingLeft: '0.3em', color: 'rgba(0, 0, 0, 0.6)', fontSize: '12px' }}>{label}</span>
                <Autocomplete
                    multiple={multiple}
                    fullWidth
                    size="small"
                    options={options}
                    getOptionLabel={(option) => option[valueField] || ""}
                    value={labelingArr[rowIndex][label] ? labelingArr[rowIndex][label] : null}
                    onChange={(event, newValue) => autoCompleteHandlerForRow(rowIndex, event, newValue, label, multiple)}
                    renderOption={(props, option) => (
                        <li {...props} key={option.id || option[valueField]}>
                            {option[valueField]}
                        </li>
                    )}
                    renderInput={(params) => (
                        <TextField {...params} variant="outlined" required={required} />
                    )}
                />
            </div>
        );
    }

    // When a field changes in a row, update that row’s data.
    function autoCompleteHandlerForRow(
        rowIndex: number,
        event: React.SyntheticEvent,
        newValue: any,
        label: string,
        multiple: boolean
    ) {
        setLabelingArr((prev) =>
            prev.map((row, index) => {
                // Only update the row that changed
                if (index !== rowIndex) return row;

                // If the WP field is changing and is different, reset related fields
                if (label === 'WP' && row.WP !== newValue) {
                    return {
                        ...row,
                        WP: newValue,
                        Phase: null,
                        "Design Stage": null,
                        Elements: null,
                        "Sub Disciplines": null,
                    };
                }

                // Otherwise, update only the specific field
                return {
                    ...row,
                    [label]: newValue,
                };
            })
        );
    }


    // COMMON FIELDS AUTOCOMPLETE (for static fields like Document Status)
    function AutoCompleteCommon(
        options: any[],
        label: string,
        valueField: string,
        required?: boolean,
        multiple: boolean = false
    ) {
        return (
            <Autocomplete
                multiple={multiple}
                fullWidth
                size="small"
                options={options}
                getOptionLabel={(option) => option[valueField] || ""}
                value={commonData[label] ? commonData[label] : null}
                onChange={(event, newValue) =>
                    setCommonData((prev: any) => ({
                        ...prev,
                        [label]: newValue,
                    }))
                }
                renderInput={(params) => (
                    <TextField {...params} label={label} variant="outlined" required={required} />
                )}
            />
        );
    }

    // Build a URL for a given row.
    function buildRowUrl(row: any) {
        const pathSegments = [
            row['Design Stage']?.Title,
            row.Elements?.ElementNameAndCode,
            row['Sub Disciplines']?.SubDiscipline,
        ].filter(Boolean);
        const rawPath = `/sites/METPRODocCenterC/${mapWP[row.WP?.Title]}/${pathSegments.join('/')}`;
        const encodedPath = encodeURI(rawPath);
        const baseUrl = `${props.context.pageContext.web.absoluteUrl}/${mapWP[row.WP?.Title]}/Forms/AllItems.aspx`;
        const url = `${baseUrl}?id=${encodedPath}`;
        return url;
    }

    // Build JSON payload for a given row, merging in commonData.
    async function buildJsonPayload(data: InputData) {

        let jsonToSave: Record<string, any> = {
            "__metadata": {
                "type": `SP.Data.${data.documentLibraryNameMapped}Item`
            }
        };

        const addField = (key: string, value: any) => {
            if (value !== null && value !== undefined && value !== "") {
                jsonToSave[key] = value;
            }
        };

        addField("Rev", data.Rev !== null && data.Rev !== undefined ? Number(data.Rev) : 0);
        addField("RevisionNo", data.RevisionNo !== null && data.Rev !== undefined ? Number(data.RevisionNo) : 0);
        addField("Authority", data.Authority)

        const addLookupField = async (key: string, lookupObject?: any) => {
            if (lookupObject !== undefined && lookupObject !== null) {
                jsonToSave[key] = {
                    "__metadata": { "type": "Collection(Edm.Int32)" },
                    "results": [lookupObject.Id]
                };
            }
        };

        await addLookupField("ElementNameAndCodeId", data.Elements);
        await addLookupField("subDisciplineId", data["Sub Disciplines"]);
        await addLookupField("OData__WPId", data.WP);
        await addLookupField("OData__designStageId", data["Design Stage"]);
        await addLookupField("OData__DocumentStatusId", data["Document Status"]);

        addField("DesignerNameId", data.AuthorDesingerName);
        addField('Phase', data.Phase?.Title)

        return jsonToSave;
    }

    // When saving, build an array of objects—one for each row—merging in the common data.
    async function saveToSP() {
        const allRowsToSave = await Promise.all(
            labelingArr.map(async (row) => {
                const rowInput: InputData = {
                    documentLibraryNameMapped: mapWP[row.WP?.Title],
                    Rev: commonData.Rev,
                    RevisionNo: commonData.RevisionNo,
                    Authority: commonData.Authority,
                    WP: row.WP,
                    Phase: row.Phase,
                    "Sub Disciplines": row["Sub Disciplines"] ? row["Sub Disciplines"] : designData.DesignDisciplinesSubDisciplines.find(e => e.Title === 'NR'),
                    Elements: row.Elements ? row.Elements : designData.Elements.find(e => e.Title === 'NR'),
                    "Design Stage": row['Design Stage'],
                    "Document Status": commonData["Document Status"],
                    AuthorDesingerName: commonData.AuthorDesingerName,
                };

                const jsonPayload = await buildJsonPayload(rowInput);

                // Build an array of URL path segments.
                // If row.Phase exists, use its Title; otherwise, it will be filtered out.
                const pathSegments = [
                    row.Phase?.Title,                // Phase (if available)
                    row['Design Stage']?.Title,      // Design Stage
                    row.Elements?.ElementNameAndCode,// Elements
                    row['Sub Disciplines']?.SubDiscipline // Sub Disciplines
                ].filter(segment => segment && segment.trim() !== '');

                // Join the segments with '/'
                const path = pathSegments.join('/');

                return {
                    ...row,
                    id: uuidv4(),
                    Rev: commonData.Rev,
                    RevisionNo: commonData.RevisionNo,
                    Authority: commonData.Authority,
                    _designType: commonData._designType,
                    "Document Status": commonData["Document Status"],
                    Elements: row.Elements ? row.Elements : designData.Elements.find(e => e.Title === 'NR'),
                    "Sub Disciplines": row["Sub Disciplines"] ? row["Sub Disciplines"] : designData.DesignDisciplinesSubDisciplines.find(e => e.Title === 'NR'),
                    AuthorDesingerName: commonData.AuthorDesingerName,
                    libraryPath: `${mapWP[row.WP?.Title]}/${path}`,
                    libraryName: `${row.WP?.Title}/${path}`,
                    documentLibraryName: row.WP?.Title,
                    documentLibraryNameMapped: mapWP[row.WP?.Title],
                    Phase: row.Phase?.Title,
                    jsonPayload: jsonPayload,
                };
            })
        );

        props.onSave(allRowsToSave);
        props.onClose();
    }

    // Handle changes to the Rev value (common field)
    function handleRevChange(event: any, label: string) {
        let value = Number(event.target.value);
        if (value > 10) {
            value = 10;
        } else if (value < 0) {
            value = 0;
        }
        setCommonData((prev: any) => ({ ...prev, [label]: value }));
    }

    // Compute whether all rows and common fields are valid.
    const allRowsValid = labelingArr.every((row) =>
        row.WP &&
        row["Design Stage"] &&
        (HAS_NO_ELEMETNS.includes(row.WP?.Title) ? true : row.Elements) &&
        (HAS_NO_SUB_DICIPLINES.includes(row.WP?.Title) ? true : row["Sub Disciplines"]) &&
        (
            !WP_PHASE_ARRAY.includes(row.WP?.Title) ||
            (row.Phase)
        )
    );

    return (
        <>
            {/* Repeating labeling rows */}
            {labelingArr.map((row, index) => {

                return (
                    <div key={uuidv4()} style={{ display: 'flex' }}>

                        <div key={row.id} className={styles.labelingContainer} style={{ width: '95%' }}>
                            {/* WP */}
                            {AutoCompleteLabeling(index, designData.Design_WP, 'WP', 'Title', true, false)}
                            {/* Phase */}
                            {AutoCompleteLabeling(
                                index,
                                row.WP
                                    ? designData.Phase.filter((phase) =>
                                        (row.WP.Title.startsWith('Wp')
                                            || WP_PHASE_ARRAY.includes(row.WP.Title)) ? phase.WP_Type === 'ALL-WP' : phase.WP_Type === 'General')
                                    : [],
                                'Phase', 'Title', true, false)}
                            {/* Design Stage */}
                            {AutoCompleteLabeling(
                                index,
                                row.WP && row.Phase
                                    ? designData.Design_DesignStage.filter((ds) =>
                                        (row.WP.Title.startsWith('Wp')
                                            || WP_PHASE_ARRAY.includes(row.WP.Title)) ? ds.WP_Type === 'ALL-WP' : ds.WP_Type === 'General' && ds.Phase.includes(row.Phase.Title))
                                    : [],
                                'Design Stage', 'Title', true, false)}
                            {/* Elements */}
                            {!HAS_NO_ELEMETNS.includes(row.WP?.Title) &&
                                AutoCompleteLabeling(
                                    index,
                                    row.WP && row.Phase && row["Design Stage"]
                                        ? designData.Elements.filter((element) => element.WP === row.WP?.Title)
                                        : [],
                                    'Elements',
                                    'ElementNameAndCode',
                                    true,
                                    false
                                )}
                            {/* Sub Disciplines */}
                            {!HAS_NO_SUB_DICIPLINES.includes(row.WP?.Title) && AutoCompleteLabeling(index, designData.DesignDisciplinesSubDisciplines, 'Sub Disciplines', 'SubDiscipline', true, false)}
                        </div>
                        <IconButton disabled={labelingArr.length === 1} size="small" sx={{ display: "flex", justifyContent: "center", width: '5%' }} onClick={() => deleteRow(index + 1)}>
                            <DeleteIcon />
                        </IconButton>
                    </div>
                );
            })}

            {/*Add new row button*/}
            <div style={{ display: 'flex', justifyContent: 'center', padding: '1em' }}>
                <Fab size="small" aria-label="add" color='success' sx={{ backgroundColor: '#8AC693' }} onClick={addRow}>
                    <AddIcon htmlColor="white" />
                </Fab>
            </div>

            {/* Static / Common Fields */}
            <div className={styles.staticContainer}>
                {/* Rev */}
                <TextField
                    value={commonData.Rev !== null && commonData.Rev !== undefined ? commonData.Rev : 0}
                    type='number'
                    label='Rev'
                    size='small'
                    fullWidth
                    onChange={(event) => handleRevChange(event, 'Rev')}
                />
                {/* Authority */}
                <TextField
                    value={commonData.Authority || ''}
                    type='text'
                    label='Authority'
                    size='small'
                    fullWidth
                    onChange={(event) => {
                        setCommonData((prev: any) => ({
                            ...prev,
                            Authority: event.target.value
                        }));
                    }}
                />
                {/* Document  Status */}
                {AutoCompleteCommon(designData.Design_DocumentStatus, 'Document Status', 'Title', false, false)}
                {/* Revision */}
                <TextField
                    value={commonData.RevisionNo !== null && commonData.RevisionNo !== undefined ? commonData.RevisionNo : 0}
                    type='number'
                    label='Revision'
                    size='small'
                    fullWidth
                    onChange={(event) => handleRevChange(event, 'RevisionNo')}
                />
                {/* Author/Designer Name */}
                <UnifiedNameAutocomplete
                    value={props.users.filter((user) => user.Id === commonData.AuthorDesingerName)[0]?.Title ?? ''}
                    size="small"
                    context={props.context}
                    users={props.users.filter((u: any) => u?.Email)}
                    multiple={false}
                    label="Author/Designer Name"
                    onChange={(idOrValue, newValue, email) => {
                        const selectedUser = props.users.find((user) => user.Title === newValue);
                        setCommonData((prev: any) => ({
                            ...prev,
                            AuthorDesingerName: selectedUser ? selectedUser.Id : null, // Store the ID instead of the name
                        }));
                    }}
                />
            </div>

            <div style={{ display: 'flex', justifyContent: 'center', marginTop: '1rem', gap: '1rem' }}>
                <Button
                    disabled={!allRowsValid}
                    style={{ textTransform: 'capitalize' }}
                    size='small'
                    variant='contained'
                    onClick={saveToSP}
                >
                    {props.dir ? 'שמור' : 'Save'}
                </Button>
                <Button
                    style={{ textTransform: 'capitalize' }}
                    size='small'
                    variant='contained'
                    color='error'
                    onClick={props.onClose}
                >
                    {props.dir ? 'בטל' : 'Cancel'}
                </Button>
            </div>
        </>
    );
}
