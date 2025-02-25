import { v4 as uuidv4 } from 'uuid';
import { BaseEntity, Entity, SchemaType, Task } from './Interfaces';
import styles from './MeetingSummaries.module.scss';
import Swal from 'sweetalert2'
import { blue, red } from '@mui/material/colors';
import moment, { Moment } from 'moment';
import { SPFI } from '@pnp/sp';

const customClass = {
    title: styles.swal2Title,
    htmlContainer: styles.swal2Content,
    confirmButton: styles.swal2Confirm,
    cancelButton: styles.swal2Cancel
}

const meetingRoomsListUrl = 'https://metprolsz.sharepoint.com/sites/METPRODocCenterC/Lists/MeetingSummaries/AllItems.aspx'

export const defaultFactory = <T>(type: SchemaType): Partial<T> => {
    const templates: Record<SchemaType, Partial<any>> = {
        Task: {
            company: '',
            department: '',
            name: '',
            designation: '',
            subject: '',
            startDate: '',
            endDate: '',
            importance: '',
            description: ''
        },
        Employee: {
            name: '',
            company: '',
            designation: '',
        },
        MeetingContent: {
            description: '',
            name: '',
            dueDate: '',
            status: ''
        }
    };

    return templates[type] as Partial<T>;
};

export const addRow = <T>(
    dataArrayName: string,
    type: SchemaType,
    setState: (updater: (prevState: any) => any) => void
): void => {
    const defaultRow = defaultFactory<T>(type);

    setState((prevState) => ({
        [dataArrayName]: [
            ...prevState[dataArrayName],
            {
                id: prevState[dataArrayName].length + 1, // Increment ID based on current length
                ...defaultRow,
                uid: uuidv4(), // Generate unique identifier
            },
        ],
    }));
};

export const deleteRow = <T extends BaseEntity>(
    dataArrayName: string,
    rowIndex: number,
    setState: (updater: (prevState: any) => any) => void
): void => {
    setState((prevState) => {
        const currentArray = prevState[dataArrayName] as T[];

        if (!currentArray || currentArray.length === 0 || rowIndex < 0 || rowIndex >= currentArray.length) {
            return prevState; // No changes if invalid index or empty array
        }

        const filteredArray = currentArray.filter((_, index) => index !== rowIndex);

        const reorderedList = filteredArray.map((item, index) => ({
            ...item,
            id: index + 1, // Reset IDs starting from 1
        }));

        return {
            [dataArrayName]: reorderedList,
        };
    });
};


export const reformatList = <T extends object>(
    list: T[],
    requiredFields: string[],
    additionalFieldsTransform?: (item: T) => Partial<T>
): (T & { id: number })[] => {
    return list
        .filter((item) =>
            requiredFields.some((field) => {
                const value = (item as any)[field];
                return value !== '' && value !== null && value !== undefined;
            })
        )
        .map((item, index) => {
            const transformedItem = {
                ...item,
                ...(additionalFieldsTransform ? additionalFieldsTransform(item) : {}),
                id: index + 1, // Assign sequential IDs
            };

            // Convert 'name' field from array to comma-separated string if it exists and is an array
            if (Array.isArray((item as any).name)) {
                (transformedItem as any).name = (item as any).name.join(', ');
            }

            return transformedItem;
        });
};

export const reformatListWithDates = <T extends object>(
    list: T[],
    dateFields: string[]
): (T & Record<string, any>)[] => {
    return list.map((item: any) => {
        const transformedItem: any = { ...item };

        dateFields.forEach((field) => {
            if (item[field]) {
                transformedItem[field] = moment(item[field]).toDate();
                transformedItem[`${field}Moment`] = moment(item[field]);
            } else {
                transformedItem[field] = null;
                transformedItem[`${field}Moment`] = null;
            }
        });

        return transformedItem;
    });
};


export const removingBlanks = <T extends object>(list: any[], requiredFields: string[]) => {
    return list
        .filter((item: T[]) => {
            // Check if at least one of the required fields is not empty/null
            return requiredFields.some((field) => {
                const value = (item as any)[field];
                return value !== '' && value !== null && value !== undefined;
            });
        })
}

export const initReformatList = <T extends object>(
    list: T[]
): (T & Record<string, any>)[] => {
    return list.map((item: any) => {
        // Convert the joined name field back to an array
        if (item.name && typeof item.name === 'string') {
            item.name = item.name.split(', ');
        }
        return item;
    });
};

export const initReformatListWithDates = <T extends object>(
    list: T[],
    dateFields: string[]
): (T & Record<string, any>)[] => {
    return list.map((item: any) => {
        const transformedItem: any = { ...item };

        dateFields.forEach((field) => {
            if (item[field]) {
                transformedItem[field] = moment(item[field]);
            } else {
                transformedItem[field] = null;
                transformedItem[`${field}Moment`] = null;
            }
        });

        return transformedItem;
    });
}

export const saveEntity = async (name: string, sp: SPFI, listId: string) => {
    try {
        await sp.web.lists.getById(listId).items.add({
            Title: name,
        });
    } catch (error) {
        console.error(`Error saving entity (${name}):`, error);
    }
};

export const saveEntities = async (
    entities: Entity[],
    sp: SPFI,
    listId: string,
    key: keyof Entity,
    ...arrays: Entity[][]
): Promise<void> => {

    console.log("key:", key)

    const combinedNames = new Set<string>(
        arrays
            .reduce((acc, array) => acc.concat(array), []) // Flatten all arrays into a single array
            .map((item: Entity) => item[key] as string)
            .filter((value: string) => value) // Filter out falsy values
    );
    let entitiesToSave: any[] = []
    if (key === 'company') {
        entitiesToSave = Array.from(new Set(combinedNames)).flat().filter(
            (value: string) => !entities.find((entity: Entity) => entity === value)
        );
    }
    else if (key === 'name') {
        entitiesToSave = Array.from(new Set(combinedNames)).flat().filter(
            (value: string) => !entities.find((entity: Entity) => entity.Title === value)
        );
    }

    if (entitiesToSave.length > 0) {
        try {
            await Promise.all(entitiesToSave.map((value: string) => saveEntity(value, sp, listId)));
        } catch (error) {
            console.error(`Error saving new ${key}s:`, error);
        }
    }
};


export const confirmSaveAndSend = async (options: any) => {

    const {
        onConfirm,
        onCancel,
        currDir
    } = options;

    const t = currDir ? require('../../../locales/he/common.json') : require('../../../locales/en/common.json') // Translator between en/he

    return Swal.fire({
        title: t.titleSaveAndSend,
        icon: "warning",
        text: t.textSaveAndSend,
        confirmButtonText: t.confirmButtonTextSaveAndSend,
        confirmButtonColor: blue.A400,
        cancelButtonText: t.No,
        cancelButtonColor: red.A700,
        showCancelButton: true,
        customClass: customClass,
        backdrop: false,
        returnFocus: false
    }).then(async (result) => {
        if (result.isConfirmed) {
            if (onConfirm) {
                await onConfirm(); // Execute the confirm callback
                window.location.href = meetingRoomsListUrl;
            }
        } else {
            if (onCancel) {
                onCancel(); // Execute the cancel callback
            }
        }
        return result.isConfirmed;
    });
};

export const sweetAlertMsgHandler = async (status: string, currDir: boolean): Promise<boolean> => {
    const t = currDir
        ? require('../../../locales/he/common.json')
        : require('../../../locales/en/common.json');

    let result;
    let timer = 1000

    if (status === "Submit") {
        Swal.fire({
            title: t.swalTitleSubmit,
            icon: "success",
            confirmButtonColor: blue.A400,
            customClass: customClass,
            willClose: () => {
                window.location.href = meetingRoomsListUrl;
            }
        }).then((confirmation) => {
            if (confirmation.isConfirmed) {
                window.location.href = meetingRoomsListUrl
            }
        })
    }

    if (status === 'send') {
        Swal.fire({
            title: t.titleSaveAndSend,
            icon: "warning",
            text: t.textSaveAndSend,
            confirmButtonText: t.Yes,
            confirmButtonColor: blue.A400,
            cancelButtonText: t.No,
            cancelButtonColor: red.A700,
            customClass: customClass,
            showCancelButton: true,
            backdrop: false,
            returnFocus: false,
            willClose: () => {
                window.location.href = meetingRoomsListUrl;
            }
        }).then((confirmation) => {
            if (confirmation.isConfirmed) {
                window.location.href = meetingRoomsListUrl;
            }
        });
    }

    if (status === 'Cancel') {
        Swal.fire({
            title: t.swalCancel,
            text: t.swalTextCancel,
            icon: "warning",
            confirmButtonText: t.Yes,
            confirmButtonColor: blue.A400,
            cancelButtonText: t.No,
            cancelButtonColor: red.A700,
            customClass: customClass,
            showCancelButton: true,
            backdrop: false,
            returnFocus: false
        }).then((confirmation) => {
            if (confirmation.isConfirmed) {
                window.location.href = meetingRoomsListUrl;
            }
        });
    }

    if (status === "SendToMeAsEmail") {
        result = await Swal.fire({
            title: t.swalTitleSendToMeAsEmail,
            icon: "success",
            confirmButtonColor: blue.A400,
            showCancelButton: true,
            confirmButtonText: t.Yes,
            cancelButtonText: t.No,
            customClass: customClass,
        });
        if (result.isConfirmed) {
            setTimeout(() => {
                window.location.href = meetingRoomsListUrl;
            }, timer);
            return true;
        } else {
            return false;
        }
    }

    if (status === "DownloadAsDraft") {
        result = await Swal.fire({
            title: t.swalTitleDownloadAsDraft,
            icon: "success",
            confirmButtonColor: blue.A400,
            showCancelButton: true,
            confirmButtonText: t.Yes,
            cancelButtonText: t.No,
            customClass: customClass,
        });
        if (result.isConfirmed) {
            setTimeout(() => {
                window.location.href = meetingRoomsListUrl;
            }, timer);
            return true;
        } else {
            return false;
        }
    }

    // If no matching status is found, do nothing and return false.
    return false;
};
