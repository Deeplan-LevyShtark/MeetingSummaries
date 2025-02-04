import React, { useEffect, useState } from 'react';
import { Autocomplete, Button, TextField } from '@mui/material';
import { SPFI } from '@pnp/sp';
import styles from './Labeling.module.scss';
import { v4 as uuidv4 } from 'uuid';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { UnifiedNameAutocomplete } from '../UnifiedNameAutocomplete/UnifiedNameAutocomplete.cmp';

interface LookupField {
    Id?: number | null;
}

interface InputData {
    documentLibraryNameMapped: string;
    Rev?: number | null;
    WP?: LookupField;
    "Sub Disciplines"?: LookupField;
    Elements?: LookupField;
    "Design Stage"?: LookupField;
    "Document Status"?: LookupField;
    AuthorDesingerName?: string | null;
}

export interface LabelingProps {
    sp: SPFI;
    context: WebPartContext;
    dir: boolean;
    users: any[];
    selectedLabeling?: any;
    onSave: (selectedLabeling: string) => void;
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
    "coordination (danon)": "cordinationDanonNew",
}

export function Labeling(props: LabelingProps) {
    // State management for labeling data    
    const [designData, setDesignData] = useState({
        Design_WP: [] as any[],
        Design_DesignStage: [] as any[],
        Elements: [] as any[],
        DesignDisciplinesSubDisciplines: [] as any[],
        Design_DocumentStatus: [] as any[],
    });

    const [selectedObject, setSelectedObject] = useState<any>(props.selectedLabeling || {});

    // Fetch labeling data on mount
    useEffect(() => {
        getLabelingData();
    }, []);

    useEffect(() => {
        if (props.selectedLabeling !== undefined) {
            setSelectedObject((prev: any) => ({
                ...prev,
                WP: props.selectedLabeling?.WP ?? null,
                "Sub Disciplines": props.selectedLabeling?.["Sub Disciplines"] ?? null,
                Elements: props.selectedLabeling?.Elements ?? null,
                "Design Stage": props.selectedLabeling?.["Design Stage"] ?? null,
                "Document Status": props.selectedLabeling?.["Document Status"] ?? null,
            }));
        }
    }, [props.selectedLabeling]);

    const getLabelingData = async () => {

        try {
            const [wp, designStage, elements, disciplines, designDocumentStatus, designType] = await Promise.all([
                props.sp.web.lists.getByTitle('Design_WP').items.select('Title, Id').top(5000)(),
                props.sp.web.lists.getByTitle('Design_DesignStage').items.top(5000)(),
                props.sp.web.lists.getByTitle('Elements').items.top(5000)(),
                props.sp.web.lists.getByTitle('DesignDisciplinesSubDisciplines').items.top(5000)(),
                props.sp.web.lists.getByTitle('Design_DocumentStatus').items.top(5000)(),
                props.sp.web.lists.getByTitle('Design_TYPE').items.top(5000)()
            ]);

            // Update state with fetched data
            setDesignData({
                Design_WP: wp,
                Design_DesignStage: designStage,
                Elements: elements,
                DesignDisciplinesSubDisciplines: disciplines,
                Design_DocumentStatus: designDocumentStatus,
            });

        } catch (error) {
            console.error('Error fetching labeling data:', error);
        }
    };

    function AutoCompleteLabeling(
        options: any[],
        label: string, // Ensure it matches selectedObject keys
        valueField: any,
        required?: boolean,
        multiple?: boolean
    ) {

        
        return (
            <Autocomplete
               multiple={multiple}
                fullWidth
                size="small"
                options={options}
                getOptionLabel={(option) => option[valueField] || ""}

                value={
                    selectedObject[label]
                      ? selectedObject[label]
                      : 
                      multiple === true
                      ? []
                      : null
                  }
                     onChange={(event, newValue) =>   
                    setSelectedObject((prev: any) => ({
                        ...prev,
                        [label]: newValue || undefined, // Ensure it updates correctly
                    }))                   
                }
                renderInput={(params) => (
                    <TextField {...params} label={label} variant="outlined" required={required} />
                )}
            />
        );
    }


    function urlBuilder() {
        // Raw path segments (do not encode here)
        const pathSegments = [
            selectedObject['Design Stage']?.Title,
            selectedObject.Elements?.ElementNameAndCode,
            selectedObject['Sub Disciplines']?.SubDiscipline,
        ].filter(Boolean); // Remove null or undefined segments

        // Join path segments with "/" without double encoding
        const rawPath = `/sites/METPRODocCenterC/${mapWP[selectedObject.WP?.Title]}/${pathSegments.join('/')}`;
        const encodedPath = encodeURI(rawPath); // Use encodeURI to encode the full path minimally

        // Base URL
        const baseUrl = `${props.context.pageContext.web.absoluteUrl}/${mapWP[selectedObject.WP?.Title]}/Forms/AllItems.aspx`;

        // Construct the full URL with encoded query parameter
        const url = `${baseUrl}?id=${encodedPath}`;

        return url;
    }

    async function buildJsonPayload(data: InputData) {
        let jsonToSave: Record<string, any> = {
            "__metadata": {
                "type": `SP.Data.${data.documentLibraryNameMapped}Item`
            }
        };

        // Helper function to add fields only if they are not null, undefined, or empty
        const addField = (key: string, value: any) => {
            if (value !== null && value !== undefined && value !== "") {
                jsonToSave[key] = value;
            }
        };

        // Ensure Rev is set, default to 0 if null
        addField("Rev", data.Rev !== null && data.Rev !== undefined ? Number(data.Rev) : 0);

        // Lookup fields (ensure Collection(Edm.Int32) format)
        const addLookupField = async (key: string, lookupObject?: any) => {
            if (lookupObject !== undefined && lookupObject !== null) {
                console.log(lookupObject);
                console.log(key);

                
                jsonToSave[key] = {
                    "__metadata": { "type": "Collection(Edm.Int32)" },
                    "results": lookupObject instanceof Array
                        ? await Promise.all(lookupObject.map((item: any) => item.Id))
                        : [ lookupObject.Id ]
                };
            }
        };
        
        console.log( data.WP);
        console.log( data["Sub Disciplines"]);
        
         await addLookupField("OData__WPId", data.WP);
         await addLookupField("subDisciplineId", data["Sub Disciplines"]);
         await addLookupField("ElementNameAndCodeId", data.Elements);
         await addLookupField("OData__designStageId", data["Design Stage"]);
         await addLookupField("OData__DocumentStatusId", data["Document Status"]);

        // String field (only add if not empty)
        addField("DesignerNameId", data.AuthorDesingerName);
        console.log(jsonToSave);
        
        return jsonToSave;
    }

    async function saveToSP() {
        // Save data to SP here
        const libraryPath = urlBuilder();

        const inputDate: InputData = {
            documentLibraryNameMapped: mapWP[selectedObject.WP[0]?.Title],
            Rev: selectedObject.Rev,
            WP: selectedObject.WP,
            "Sub Disciplines": selectedObject["Sub Disciplines"],
            Elements: selectedObject?.Elements,
            "Design Stage": selectedObject['Design Stage'],
            "Document Status": selectedObject['Document Status'],
            AuthorDesingerName: selectedObject?.AuthorDesingerName
        };

        const jsonPayload = await buildJsonPayload(inputDate);
        console.log(selectedObject["Sub Disciplines"][0].SubDiscipline);
        

        const selectedLabeling = {
            ...selectedObject,
            Id: uuidv4(),
            libraryPath: `${mapWP[selectedObject?.WP[0]?.Title]}/${selectedObject['Design Stage'][0]?.Title}/${selectedObject.Elements[0]?.ElementNameAndCode}/${selectedObject['Sub Disciplines'][0]?.SubDiscipline}`,
            libraryName: `${selectedObject?.WP[0]?.Title}/${selectedObject['Design Stage'][0]?.Title}/${selectedObject.Elements[0]?.ElementNameAndCode}/${selectedObject['Sub Disciplines'][0]?.SubDiscipline}`,
            documentLibraryName: selectedObject?.WP.Title,
            documentLibraryNameMapped: mapWP[selectedObject?.WP.Title],
            jsonPayload: jsonPayload,
        }
        props.onSave(selectedLabeling);
        props.onClose();
    }

    function handleRevChange(event: any, name: string) {
        // only values from 1-10
        if (event.target.value > 10) {
            event.target.value = 10;
        } else if (event.target.value < 1) {
            event.target.value = 1;
        }
        setSelectedObject({ ...selectedObject, [name]: event.target.value });
    }

    return (
        <>
        
            <div className={styles.labelingContainer}>
        {console.log(selectedObject.WP)}
                {AutoCompleteLabeling(designData.Design_WP, 'WP', 'Title', true,true)}
                {AutoCompleteLabeling(designData.Design_DesignStage, 'Design Stage', 'Title', true,true)}
                {AutoCompleteLabeling(selectedObject.WP ? designData.Elements.filter((element) => element.WP === selectedObject.WP[0]?.Title) : designData.Elements,
                    'Elements', 'ElementNameAndCode', true,true)}
                {AutoCompleteLabeling(designData.DesignDisciplinesSubDisciplines, 'Sub Disciplines', 'SubDiscipline', true,true)}
                <TextField value={selectedObject.Rev !== null && selectedObject.Rev !== undefined ? selectedObject.Rev : 0} type='number' label='Rev' size='small' fullWidth onChange={(event) => handleRevChange(event, 'Rev')}></TextField>
                {AutoCompleteLabeling(designData.Design_DocumentStatus, 'Document Status', 'Title',false,true)}
                <UnifiedNameAutocomplete
                    value={
                        props.users.filter((user) => user.Id === selectedObject.AuthorDesingerName)[0]?.Title ?? ''
                    }
                    size="small"
                    context={props.context}
                    users={props.users.filter((u: any) => u?.Email)}
                    multiple={false}
                    label="Author/Designer Name"
                    onChange={(idOrValue, newValue, email) => {
                        const selectedUser = props.users.find((user) => user.Title === newValue);
                        setSelectedObject((prev: any) => ({
                            ...prev,
                            AuthorDesingerName: selectedUser ? selectedUser.Id : null, // Store the ID instead of the name
                        }));
                    }}
                />

            </div>
            <div style={{ display: 'flex', justifyContent: 'center', marginTop: '1rem', gap: '1rem' }}>
                <Button
                    disabled={!selectedObject.WP || !selectedObject['Design Stage'] || !selectedObject.Elements || !selectedObject['Sub Disciplines']}
                    style={{ textTransform: 'capitalize' }}
                    size='small'
                    variant='contained'
                    onClick={() => saveToSP()}
                >
                    {props.dir ? 'שמור' : 'Save'}
                </Button>
                <Button
                    style={{ textTransform: 'capitalize' }}
                    size='small'
                    variant='contained'
                    color='error'
                    onClick={() => props.onClose()}
                >
                    {props.dir ? 'בטל' : 'Cancel'}
                </Button>
            </div>
        </>
    );
}
