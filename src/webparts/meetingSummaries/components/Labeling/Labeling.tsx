import React, { useEffect, useState } from 'react';
import { Autocomplete, Button, TextField } from '@mui/material';
import { SPFI } from '@pnp/sp';
import styles from './Labeling.module.scss';
import { v4 as uuidv4 } from 'uuid';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { UnifiedNameAutocomplete } from '../UnifiedNameAutocomplete/UnifiedNameAutocomplete.cmp';

export interface LabelingProps {
    sp: SPFI;
    context: WebPartContext;
    dir: boolean;
    users: any[];
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
    const [Design_WP, setDesign_WP] = useState<any[]>([]);
    const [Design_DesignStage, setDesign_DesignStage] = useState<any[]>([]);
    const [Elements, setElements] = useState<any[]>([]);
    const [DesignDisciplinesSubDisciplines, setDesignDisciplinesSubDisciplines] = useState<any[]>([]);
    const [Design_DocumentStatus, setDesign_DocumentStatus] = useState<any[]>([]);
    const [Design_TYPE, setDesign_TYPE] = useState<any[]>([]);

    const [selectedObject, setSelectedObject] = useState<any>({});

    // Fetch labeling data on mount
    useEffect(() => {
        getLabelingData();
    }, []);

    useEffect(() => {
        console.log('Selected Object:', selectedObject);
    }, [selectedObject]);

    const getLabelingData = async () => {
        console.log(props.context.pageContext.web.absoluteUrl);

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
            setDesign_WP(wp);
            setDesign_DesignStage(designStage);
            setElements(elements);
            setDesignDisciplinesSubDisciplines(disciplines);
            setDesign_DocumentStatus(designDocumentStatus);
            setDesign_TYPE(designType);

        } catch (error) {
            console.error('Error fetching labeling data:', error);
        }
    };

    // Generic Autocomplete component
    function AutoCompleteLabeling(options: any[], label: string, valueField: string, required?: boolean) {
        return (
            <Autocomplete
                fullWidth
                size='small'
                options={options}
                getOptionLabel={(option) => option[valueField] || ''}
                onChange={(event, newValue) => setSelectedObject({ ...selectedObject, [label]: newValue })}
                renderInput={(params) => (
                    <TextField {...params} label={label} variant="outlined" />
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

    function saveToSP() {
        // Save data to SP here
        const libraryPath = urlBuilder();

        const selectedLabeling = {
            ...selectedObject,
            Id: uuidv4(),
            libraryPath: `${mapWP[selectedObject?.WP.Title]}/${selectedObject['Design Stage']?.Title}/${selectedObject.Elements?.ElementNameAndCode}/${selectedObject['Sub Disciplines']?.SubDiscipline}`,
            libraryName: `${selectedObject?.WP.Title}/${selectedObject['Design Stage']?.Title}/${selectedObject.Elements?.ElementNameAndCode}/${selectedObject['Sub Disciplines']?.SubDiscipline}`,
            documentLibraryName: selectedObject?.WP.Title,
            documentLibraryNameMapped: mapWP[selectedObject?.WP.Title]
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

    const filteredElements = Elements.filter((element) => element.WP === selectedObject.WP?.Title);
    // const filteredDesignStage = Design_DesignStage.filter((designStage) => designStage.Design_Stage === selectedObject['Design Stage']?.Title);

    return (
        <>
            <div className={styles.labelingContainer}>
                {AutoCompleteLabeling(Design_WP, 'WP', 'Title', true)}
                {AutoCompleteLabeling(Design_DesignStage, 'Design Stage', 'Title', true)}
                {AutoCompleteLabeling(selectedObject.WP ? filteredElements : Elements, 'Elements', 'ElementNameAndCode', true)}
                {AutoCompleteLabeling(DesignDisciplinesSubDisciplines, 'Sub Disciplines', 'SubDiscipline', true)}
                <TextField type='number' label='Rev' size='small' fullWidth onChange={(event) => handleRevChange(event, 'Rev')}></TextField>
                {AutoCompleteLabeling(Design_DocumentStatus, 'Document Status', 'Title')}
                <TextField type='number' label='Revision' size='small' fullWidth onChange={(event) => handleRevChange(event, 'Revision')}></TextField>
                <UnifiedNameAutocomplete size='small' context={props.context} users={props.users} multiple={false} label='Author/Designer Name'
                    onChange={(newValue: any) => setSelectedObject({ ...selectedObject, AuthorDesingerName: newValue })} />
                {/* {AutoCompleteLabeling(Design_TYPE, 'Design Type', 'Title')} */}
            </div>
            <div style={{ display: 'flex', justifyContent: 'center', marginTop: '1rem' }}>
                <Button
                    disabled={!selectedObject.WP || !selectedObject['Design Stage'] || !selectedObject.Elements || !selectedObject['Sub Disciplines']}
                    style={{ textTransform: 'capitalize' }}
                    size='small'
                    variant='contained'
                    onClick={() => saveToSP()}
                >
                    {props.dir ? 'שמור' : 'Save'}
                </Button>
            </div>
        </>
    );
}
