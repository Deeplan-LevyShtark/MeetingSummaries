import React, { createRef } from 'react';
import styles from './MeetingSummaries.module.scss';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { TextField, Paper, Divider, LinearProgress, Button, Typography } from '@mui/material';
import { Header } from './Header/Header';
import Loader from './Loader/Loader.cmp';
import { DatePicker } from '@mui/x-date-pickers/DatePicker';
import { AdapterMoment } from '@mui/x-date-pickers/AdapterMoment';
import { LocalizationProvider } from '@mui/x-date-pickers/LocalizationProvider';
import { createTheme, ThemeProvider } from '@mui/material/styles';
import { FolderPicker, IFolder } from "@pnp/spfx-controls-react/lib/FolderPicker";
import moment, { Moment } from 'moment';
import { Schema, StyledDatePicker, Task, Employee, MeetingContent } from './Interfaces'
import { Icon } from '@fluentui/react/lib/Icon';
import ClearIcon from '@mui/icons-material/Clear';
import TableRepeatingSection from './TableReaptingSection/TableRepeatingSection.cmp';
import { IconButton } from '@mui/material';
import { CacheProviderWrapper } from './CacheProviderWrapper';
import CloseIcon from '@mui/icons-material/Close';
import { v4 as uuidv4 } from 'uuid';
import { addRow, deleteRow, sweetAlertMsgHandler, reformatList, reformatListWithDates, initReformatListWithDates, saveEntities, initReformatList, confirmSaveAndSend } from './Utils';
import PeoplePickerMUI from './PeoplePickerMUI/PeoplePickerMUI.cmp';
import PopUp from './PopUp/PopUp.cmp';
import { Labeling } from './Labeling/Labeling';
import SourceIcon from '@mui/icons-material/Source';
import { NewContact } from './NewContact/NewContact';
import PersonAddIcon from '@mui/icons-material/PersonAdd';

export interface IMeetingSummariesEditProps {
    userDisplayName: string;
    Title: string;
    MeetingSummariesListId: string;
    CompaniesList: string;
    ExternalUsersOptions: string;
    TasksListId: string;
    sp: SPFI;
    context: WebPartContext;
    FormID: string;
}

export interface IMeetingSummariesEditStates {
    currUser: any;
    LoadingForm: string;
    users: any[];
    MeetingSummary: string;
    DateOfMeeting: Moment;
    selectionModel: any[];
    libraryPath: string;
    libraryName: string;
    isValid: boolean;
    [key: string]: any;
    errors: { [key: string]: string | [] };
    currDir: boolean;
    companies: any[];
    attendees: Employee[];
    absents: Employee[];
    tasks: Task[];
    meetingContent: MeetingContent[];
    externalUsers: any[];
    externalUsersToSave: any[];
    selectedUsers: any[];
    selectedUsersFreeSolo: any[];
    freeSoloUser: string;
    submit: string;
    selectedLabeling: any;
    addNewContactPopUp: boolean;
}

export const createDirTheme = (isRtl: boolean) =>
    createTheme({
        direction: isRtl ? 'rtl' : 'ltr', // Conditionally set direction
        typography: {
            fontFamily: `'Roboto', 'Arial', sans-serif`,
        },
    });


const theme = createTheme({
    components: {
        MuiOutlinedInput: {
            styleOverrides: {
                root: {
                    borderRadius: 0, // Disable border radius for input fields
                },
            },
        },
    },
});

export default class MeetingSummariesEdit extends React.Component<IMeetingSummariesEditProps, IMeetingSummariesEditStates> {
    public apiRef: any;

    constructor(props: IMeetingSummariesEditProps) {
        super(props);
        this.apiRef = createRef() as any;
        this.state = {
            currUser: null,
            LoadingForm: 'Loading',
            users: [],
            MeetingSummary: '',
            DateOfMeeting: moment(),
            selectionModel: [],
            libraryPath: '',
            libraryName: '',
            isValid: true,
            errors: {},
            currDir: false,
            companies: [],
            attendees: [{ id: 1, name: '', company: '', designation: '', uid: uuidv4(), idView: '1' }],
            absents: [{ id: 1, name: '', company: '', designation: '', uid: uuidv4(), idView: '1' }],
            tasks: [{ id: 1, company: '', name: '', designation: '', department: '', subject: '', startDate: '', endDate: '', importance: '', description: '', uid: uuidv4(), idView: '1' }],
            meetingContent: [{ id: 1, description: '', name: '', dueDate: '', status: '', uid: uuidv4(), idView: '1' }],
            externalUsers: [],
            externalUsersToSave: [],
            selectedUsers: [],
            selectedUsersFreeSolo: [],
            freeSoloUser: '',
            submit: '',
            selectedLabeling: {},
            addNewContactPopUp: false
        }

        this.onChangeGeneric = this.onChangeGeneric.bind(this);
    }

    componentDidMount(): void {
        this.onInit()
        setTimeout(() => {
            this.setState({ LoadingForm: 'ok' })
        }, 2200);
    }

    componentDidUpdate(prevProps: Readonly<IMeetingSummariesEditProps>, prevState: Readonly<IMeetingSummariesEditStates>, snapshot?: any): void {
        // console.log('this.state', this.state);
    }

    onInit = async () => {
        try {
            // Run all promises in parallel
            const [user, companiesList, externalUsers, item, users] = await Promise.all([
                this.props.sp.web.currentUser()
                    .catch(err => { console.error("Error fetching current user:", err); throw err }),
                this.props.sp.web.lists.getById(this.props.CompaniesList).items.select('Title')()
                    .catch(err => { console.error("Error fetching Companies:", err); throw err }),
                this.props.sp.web.lists.getById(this.props.ExternalUsersOptions).items()
                    .catch(err => { console.error("Error fetching External Users:", err); throw err }),
                this.props.sp.web.lists.getById(this.props.MeetingSummariesListId).items.getById(parseInt(this.props.FormID))()
                    .catch(err => { console.error("Error fetching Meeting Summary:", err); throw err }),
                this.props.sp.web.siteUsers.select('Id, Title, Email').filter("Email ne ''")()
                    .catch(err => { console.error("Error fetching Users:", err); throw err }),
            ]);

            // Extract only the 'Title' from companies
            const companies = companiesList?.map((company) => company.Title) || [];

            // Update state with resolved data
            this.setState({
                MeetingSummary: item.MeetingSummary,
                DateOfMeeting: moment(item.DateOfMeeting),
                libraryPath: item.libraryPath,
                libraryName: item.libraryName,
                attendees: JSON.parse(item.attendees).length !== 0 ? initReformatList(JSON.parse(item.attendees)) : [{ id: 1, name: '', company: '', designation: '', uid: uuidv4() }],
                absents: JSON.parse(item.absents).length !== 0 ? initReformatList(JSON.parse(item.absents)) : [{ id: 1, name: '', company: '', designation: '', uid: uuidv4() }],
                tasks: JSON.parse(item.tasks).length !== 0 ? initReformatListWithDates(initReformatList(JSON.parse(item.tasks)), ['startDate', 'endDate']) : [{ id: 1, company: '', name: '', designation: '', department: '', subject: '', startDate: '', endDate: '', importance: '', description: '', uid: uuidv4() }],
                meetingContent: JSON.parse(item.meetingContent).length !== 0 ? initReformatListWithDates(initReformatList(JSON.parse(item.meetingContent)), ['dueDate']) : [{ id: 1, description: '', name: '', dueDate: '', status: '', uid: uuidv4() }],
                currUser: user,
                users: [...users, ...externalUsers],
                companies: companies,
                externalUsers: externalUsers,
                currDir: item.dir,
                selectedUsers: JSON.parse(item.selectedUsers),
                selectedUsersFreeSolo: JSON.parse(item.selectedUsersFreeSolo),
                submit: item.submit,
                selectedLabeling: JSON.parse(item.selectedLabeling)
            });
        } catch (error) {
            console.error("Error initializing data:", error);
        }
    };

    onChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
        const { name, value } = e.target;

        this.setState({
            [name]: value,
        } as Pick<IMeetingSummariesEditStates, keyof IMeetingSummariesEditStates>);
        if (name === 'MeetingSummary') {
            this.handleErrorRequire(value, 'MeetingSummary')
        }
    };

    handleDateChange = (value: any): void => {
        this.setState({ DateOfMeeting: value }, () => {
            this.handleErrorRequire(value, 'DateOfMeeting')
        });
    }

    folderHandle(folder: IFolder) {
        let libraryPath = folder.ServerRelativeUrl

        this.setState({ libraryPath: libraryPath, libraryName: folder.Name }, () => {
            this.handleErrorRequire(libraryPath, 'libraryPath')
        })
    }

    libraryPathHandle = (selectedLabeling: any): void => {
        this.setState({ libraryPath: selectedLabeling?.libraryPath, libraryName: selectedLabeling?.libraryName, selectedLabeling: selectedLabeling }, () => {
            this.handleErrorRequire(selectedLabeling?.libraryPath, 'libraryPath')
        })
    }

    closeFolderPopUp = (): void => {
        this.setState({ folderPopUp: false })
    }

    closeAddNewContactPopUp = async (): Promise<void> => {
        const externalUsers = await this.props.sp.web.lists.getById(this.props.ExternalUsersOptions).items()
        this.setState({ addNewContactPopUp: false, users: [...this.state.users, ...externalUsers] })
    }

    // Validation for the entire form
    validateForm = (): boolean => {
        const requiredFields = ['DateOfMeeting', 'MeetingSummary', 'libraryPath'];
        const t = this.state.currDir ? require('../../../locales/he/common.json') : require('../../../locales/en/common.json')

        const errors: { [key: string]: string | [] } = {};
        let isValid = true;

        // Check required fields
        requiredFields.forEach((field) => {
            const value = this.state[field];
            if (moment.isMoment(value) && !value.isValid()) {
                errors[field] = t.required
                isValid = false;
            }
            if (!value || value === '') {
                errors[field] = t.required;
                isValid = false;
            }
        });
        console.log("errors:", errors)

        this.setState({ errors });
        return isValid;
    };

    handleErrorRequire = (value: string | number | any[] | Object, name: string) => {
        const t = this.state.currDir ? require('../../../locales/he/common.json') : require('../../../locales/en/common.json')
        this.setState((prevState) => {
            const errors = { ...prevState.errors };

            const isValid = (val: any) => {
                if (val === null || val === '' || val === undefined) {
                    return false;
                }

                if (Array.isArray(val)) {
                    return val.length > 0;
                }

                if (typeof val === 'number') {
                    return val > 0;
                }

                if (typeof val === 'string') {
                    return val.trim().length > 0;
                }

                // Check if the value is a moment date and if it's before now
                if (moment.isMoment(val)) {
                    return val.isValid()
                }

                return true;
            };

            if (isValid(value)) {
                delete errors[name];
            } else {
                errors[name] = t.required;
            }
            return { [name]: value, errors };
        })

    }

    requireErrorJSX = (name: string): JSX.Element | null => {
        return this.state.errors[name]
            ? <div style={{ display: 'flex', width: '100%', paddingTop: '5px' }}>
                <Icon style={{ display: 'flex', alignItems: 'center', color: 'rgb(211, 47, 47)', marginRight: '5px' }} iconName='Error'></Icon>
                <span className={styles.errorMessage}>{this.state.errors[name]}</span>
            </div>
            : null;
    }

    submitForm = async (submitType: string) => {
        const { users, currDir, companies, DateOfMeeting, MeetingSummary, libraryPath, libraryName, attendees, absents, meetingContent, tasks, selectedUsers, selectedUsersFreeSolo, currUser } = this.state
        this.setState({ LoadingForm: 'Saving' })

        const reformattedAttendees = reformatList<Employee>(
            attendees,
            ['company', 'name', 'designation']
        );

        const reformattedAbsents = reformatList<Employee>(
            absents,
            ['company', 'name', 'designation']
        );

        const reformattedMeetingContent =
            reformatList<MeetingContent>(
                reformatListWithDates(meetingContent, ['dueDate']),
                ['description', 'name', 'dueDate', 'status']
            );

        const reformattedTasks =
            reformatList<Task>(
                reformatListWithDates(tasks, ['startDate', 'endDate']),
                ['company', 'name', 'designation', 'department', 'subject', 'startDate', 'endDate', 'importance', 'description', 'ids']
            )


        if (this.validateForm()) {
            if (submitType === 'save') {
                try {
                    await Promise.all([
                        // Users
                        saveEntities(users, this.props.sp, this.props.ExternalUsersOptions, 'name', attendees, absents, tasks, meetingContent),
                        // Companies
                        saveEntities(companies, this.props.sp, this.props.CompaniesList, 'company', attendees, absents, meetingContent)
                    ]);
                } catch (err) {
                    console.error("Error saving entites Users or Companies", err)
                }
                try {
                    await this.props.sp.web.lists.getById(this.props.MeetingSummariesListId).items.getById(Number(this.props.FormID)).update({
                        DateOfMeeting: moment(DateOfMeeting),
                        MeetingSummary: MeetingSummary,
                        attendees: JSON.stringify(reformattedAttendees),
                        absents: JSON.stringify(reformattedAbsents),
                        meetingContent: JSON.stringify(reformattedMeetingContent),
                        tasks: JSON.stringify(reformattedTasks),
                        libraryPath: libraryPath,
                        libraryName: libraryName,
                        selectedUsers: JSON.stringify(selectedUsers),
                        selectedUsersFreeSolo: JSON.stringify(selectedUsersFreeSolo),
                        submit: submitType,
                        Summarizing: currUser?.Title,
                        Copy: [...this.state.selectedUsers, ...this.state.selectedUsersFreeSolo].flat().join(', '),
                        FormLink: {
                            Description: MeetingSummary,
                            Url: `${this.props.context.pageContext.web.absoluteUrl}/SitePages/MeetingSummaries.aspx?FormID=${this.props.FormID}`
                        }

                    })

                } catch (error) {
                    console.error(error);
                }
                sweetAlertMsgHandler('Submit', currDir)
            }
            if (submitType === 'send') {
                // Show confirmation dialog
                await confirmSaveAndSend({
                    currDir,
                    onConfirm: async () => {
                        try {
                            await this.props.sp.web.lists.getById(this.props.MeetingSummariesListId).items.getById(Number(this.props.FormID)).update({
                                DateOfMeeting: moment(DateOfMeeting),
                                MeetingSummary: MeetingSummary,
                                attendees: JSON.stringify(reformattedAttendees),
                                absents: JSON.stringify(reformattedAbsents),
                                meetingContent: JSON.stringify(reformattedMeetingContent),
                                tasks: JSON.stringify(reformattedTasks),
                                libraryPath: libraryPath,
                                libraryName: libraryName,
                                selectedUsers: JSON.stringify(selectedUsers),
                                selectedUsersFreeSolo: JSON.stringify(selectedUsersFreeSolo),
                                submit: submitType,
                                Summarizing: currUser?.Title,
                                Copy: [...this.state.selectedUsers, ...this.state.selectedUsersFreeSolo].flat().join(', '),
                                FormLink: {
                                    Description: MeetingSummary,
                                    Url: `${this.props.context.pageContext.web.absoluteUrl}/SitePages/MeetingSummaries.aspx?FormID=${this.props.FormID}`
                                }
                            })

                        } catch (error) {
                            console.error(error);
                        }

                        for (const task of reformattedTasks) {
                            try {
                                const filteredAssignToExternal = users.filter(user =>
                                    task.name.split(', ').includes(user.Title)
                                );
                                const filterAssignToInternal = task.name
                                    .split(', ')
                                    .filter(name => filteredAssignToExternal.every(user => user.Title !== name));

                                await this.props.sp.web.lists.getById(this.props.TasksListId).items.add({
                                    Title: task.subject,
                                    MeetingSummaryDate: DateOfMeeting,
                                    AssignedToInternalId: task.ids && task.ids.filter(id => id !== ''),
                                    AssignedToExternal: task.ids && task.ids.filter(id => id !== '').length === 0 ? task.name : filterAssignToInternal.join(', '),
                                    MeetingSummaryName: MeetingSummary,
                                    StartDate: task.startDate,
                                    EndDate: task.endDate,
                                    Description: task.description,
                                    Designation: task.designation,
                                    Importance: task.importance,
                                    UUID: task.uid,
                                    Company: task.company,
                                    ForInfoId: task.forInfoIds,
                                    LinkToMeetingSummary: {
                                        Description: MeetingSummary,
                                        Url: `${this.props.context.pageContext.web.absoluteUrl}/SitePages/MeetingSummaries.aspx?FormID=${this.props.FormID}`
                                    }
                                });
                            } catch (err) {
                                console.error("Error saving task", task, err);
                            }
                        }
                        console.log('Tasks processed.');
                    }
                })
            }
        }

        this.setState({ LoadingForm: 'ok' })
    }

    // * new functions /
    onChangeGeneric(e: any, dataArrayName: string, fieldName: string, rowIndex: number, onBlur?: string): void {
        const value = e.target.value;

        let company = ''

        if (e?.target?.Email?.length > 0 && e?.target?.Email[0] !== undefined) {
            let selectedEmails = e?.target?.Email
            selectedEmails = selectedEmails?.filter((email: string) => email !== '' && email !== undefined);
            const filterdEmailsLsz = selectedEmails?.filter((email: string) => email.includes("lsz"));
            const filterdEmailsDmy = selectedEmails?.filter((email: string) => email.includes("dmy"));
            company = filterdEmailsLsz?.length === selectedEmails?.length ? "lsz" : filterdEmailsDmy?.length === selectedEmails?.length ? "dmy" : "other";
        }

        this.setState((prevState: any) => {
            const updatedArray = [...prevState[dataArrayName]];

            if (dataArrayName !== 'tasks' && dataArrayName !== 'meetingContent' && fieldName !== 'designation') {

                if (e?.target?.Email?.length > 0) {

                    if (company === 'other') {
                        updatedArray[rowIndex] = { ...updatedArray[rowIndex], company: '', [fieldName]: value };
                    }
                    else if (company === 'lsz') {
                        updatedArray[rowIndex] = { ...updatedArray[rowIndex], company: 'MetPro-LSZ', [fieldName]: value };
                    }
                    else if (company === 'dmy') {
                        updatedArray[rowIndex] = { ...updatedArray[rowIndex], company: 'MetPro-DMY', [fieldName]: value };
                    } else {
                        updatedArray[rowIndex] = {
                            ...updatedArray[rowIndex], [fieldName]: value
                        }
                    }

                } else {
                    updatedArray[rowIndex] = {
                        ...updatedArray[rowIndex], company: '', [fieldName]: value
                    }
                }
            } else {
                updatedArray[rowIndex] = { ...updatedArray[rowIndex], [fieldName]: value }
            }

            // Combine all "name" fields from attendees, absents, and tasks
            const combinedSelectedUsers = new Set([
                ...prevState.attendees.map((item: any) => item.name).filter((name: string) => name).flat(),
                ...prevState.absents.map((item: any) => item.name).filter((name: string) => name).flat(),
            ]);

            if (onBlur !== 'onBlur' && fieldName === "name") {
                updatedArray[rowIndex] = { ...updatedArray[rowIndex], ids: e.target.ids ? [...e.target.ids] : [] }
            }

            if (onBlur !== 'onBlur' && fieldName === "forInfo") {
                updatedArray[rowIndex] = { ...updatedArray[rowIndex], forInfoIds: e.target.forInfoIds ? [...e.target.forInfoIds] : [] }
            }

            return { [dataArrayName]: updatedArray, selectedUsers: Array.from(combinedSelectedUsers) };
        });
    }

    onClickFreeSolo = () => {
        const { freeSoloUser } = this.state;

        if (!freeSoloUser.trim()) {
            // If freeSoloUser is empty or only spaces, do nothing
            return;
        }

        this.setState((prevState: any) => ({
            selectedUsersFreeSolo: Array.from(
                new Set([...prevState.selectedUsersFreeSolo, freeSoloUser.trim()])
            ),
            freeSoloUser: '' // Reset the input field
        }));
    };

    deleteFreeSoloUser = (e: any, name: string) => {
        const filterdSoloUser = this.state.selectedUsersFreeSolo.filter(item => item !== name)
        this.setState({ selectedUsersFreeSolo: filterdSoloUser })
    }

    attachRow = (dataArrayName: string, rowIndex: number): void => {
        const currentArray = this.state[dataArrayName] as MeetingContent[];

        // Validate currentArray and rowIndex
        if (!Array.isArray(currentArray) || rowIndex < 0 || rowIndex >= currentArray.length) {
            console.error("Invalid dataArrayName or rowIndex");
            return;
        }

        const { description, name, dueDate, uid, ids } = currentArray[rowIndex];
        console.log('len', this.state.tasks.length)
        const taskTemplate: Task = {
            id: this.state.tasks.length === 1
                && !this.state.tasks[0].name
                && !this.state.tasks[0].description
                && !this.state.tasks[0].endDate // Check if dueDate is also empty
                ? 1 // Keep the ID as 1 for the first empty row
                : this.state.tasks.length + 1, // Increment for subsequent tasks

            company: '',
            name: name || '', // Default to empty string
            designation: '',
            department: '',
            subject: '',
            startDate: '',
            ids: ids || [], // Ensure valid ids
            endDate: dueDate || '', // Ensure valid date
            importance: '',
            description: description || '',
            uid: uuidv4(),
            attachedId: uid,
        };

        this.setState((prevState) => {
            const tasks = Array.isArray(prevState.tasks) ? [...prevState.tasks] : [];

            // Prevent duplicate attachments
            const isDuplicate = tasks.some((task) => task.attachedId === uid);
            if (isDuplicate) {
                console.warn("Task already attached:", uid);
                return null; // No update to state
            }

            // Handle first row empty case
            if (tasks.length > 0 && !tasks[0].name && !tasks[0].description && !tasks[0].attachedId) {
                tasks[0] = taskTemplate; // Overwrite the empty first row
            } else {
                tasks.push(taskTemplate); // Append to the tasks array
            }

            return { tasks };
        });
    };
    public render(): React.ReactElement<IMeetingSummariesEditProps> {

        const { currUser, currDir, LoadingForm, DateOfMeeting, MeetingSummary, users, libraryName, errors, attendees, absents, tasks, meetingContent } = this.state

        const t = currDir ? require('../../../locales/he/common.json') : require('../../../locales/en/common.json') // Translator between en/he

        const employeeSchema: Schema = {
            fields: [
                { name: 'id', label: t.ID, type: 'number', width: 50 },
                { name: 'name', label: t.Name, type: 'string', editable: true },
                { name: 'company', label: t.Company, type: 'string', width: 180, editable: true },
                { name: 'designation', label: t.Designation, type: 'string', editable: true },
                { name: 'action', label: t.Delete, type: 'action', width: 50, editable: true },
            ],
        }
        const taskSchema: Schema = {
            fields: [
                // { name: 'id', label: t.ID, type: 'number', width: 80 },
                { name: 'company', label: t.Company, type: 'string', editable: true },
                { name: 'name', label: t.Name, type: 'string', editable: true },
                { name: 'designation', label: t.Designation, type: 'string', editable: true },
                { name: 'department', label: t.Department, type: 'string', editable: true },
                { name: 'subject', label: t.Subject, type: 'string', editable: true },
                { name: 'startDate', label: t['Start Date'], type: 'date', editable: true },
                { name: 'endDate', label: t['End Date'], type: 'date', editable: true },
                { name: 'importance', label: t.Importance, type: 'string', editable: true },
                { name: 'description', label: t.Description, type: 'string', editable: true },
                { name: 'forInfo', label: t.forInfo, type: 'string', editable: true },
                { name: 'action', label: t.Delete, type: 'action', width: 50, editable: true }
            ],
        }
        const meetingContentSchema: Schema = {
            fields: [
                { name: 'idView', label: t.ID, type: 'number', width: 80 },
                { name: 'description', label: t.Description, type: 'string', editable: true },
                { name: 'name', label: t.ActionBy, type: 'string', width: 180, editable: true },
                { name: 'dueDate', label: t.DueDate, type: 'date', width: 180, editable: true },
                { name: 'status', label: t.Status, type: 'string', width: 180, editable: true },
                { name: 'attach', label: t.Attach, type: 'action', width: 50, editable: true },
                { name: 'action', label: t.Delete, type: 'action', width: 50, editable: true },
            ]
        }


        return (
            <LocalizationProvider dateAdapter={AdapterMoment} adapterLocale={currDir ? 'he' : 'en-gb'}>

                <CacheProviderWrapper isRtl={currDir}>
                    <form dir={currDir ? 'rtl' : 'ltr'} style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>

                        <Paper elevation={3} className={styles.paperContainer}>
                            <Header context={this.props.context} Title={t.Title}></Header>
                            <div className={styles.ContainerForm}>

                                {LoadingForm === 'Loading' ? <Loader /> : <section>
                                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', paddingBottom: '1em' }}>
                                        <div style={{ display: 'flex', justifyContent: 'start', gap: '1em', paddingTop: '1em', paddingBottom: '1em' }}>
                                            <Button
                                                variant="contained"
                                                size="medium"
                                                startIcon={<PersonAddIcon />} // Icon on the left side
                                                onClick={() => this.setState({ addNewContactPopUp: true })}
                                                sx={{ display: "flex", gap: "0.5em", textTransform: "capitalize" }}
                                            >
                                                {this.state.currDir ? "איש קשר חדש" : "Add New Contact"}
                                            </Button>
                                        </div>
                                    </div>

                                    <section className={styles.Section}>
                                        <div className={styles.fieldStyle}>
                                            <DatePicker onChange={(value: any) => this.handleDateChange(value)} sx={{ width: '100%' }} disabled
                                                format='DD/MM/YYYY' label={t.Date} value={moment()}
                                            ></DatePicker>
                                        </div>

                                        <div className={styles.fieldStyle}>
                                            <TextField type='text' onChange={this.onChange} name='MeetingSummary'
                                                fullWidth label={t["Meeting Summary"]} value={MeetingSummary} variant="outlined" required={true} error={!!errors['MeetingSummary']} />
                                            {this.requireErrorJSX('MeetingSummary')}
                                        </div>

                                        <div className={styles.fieldStyle}>
                                            <StyledDatePicker onChange={(value: any) => this.handleDateChange(value)} sx={{ width: '100%' }}
                                                format='DD/MM/YYYY' label={`${t["Date of meeting"]} *`} error={!!errors['DateOfMeeting']} value={DateOfMeeting}
                                            ></StyledDatePicker>
                                            {this.requireErrorJSX('DateOfMeeting')}
                                        </div>

                                        <PeoplePickerMUI
                                            context={this.props.context}
                                            users={users}
                                            selectedUser={currUser}
                                            label={t.Summarizing}
                                            multiple={false}
                                            onChange={(newValue) => {
                                                this.setState({
                                                    currUser: newValue ? { Title: newValue.Title, Email: newValue.Email } : { Title: "", Email: "" },
                                                });
                                            }}
                                        />

                                    </section>

                                    <PopUp open={this.state.addNewContactPopUp} title={this.state.currDir ? "איש קשר חדש" : 'Add New Contact'} onClose={() => this.closeAddNewContactPopUp()} actions={null} dir={currDir ? 'rtl' : 'ltr'}>
                                        <NewContact onClose={this.closeAddNewContactPopUp} dir={currDir} sp={this.props.sp} context={this.props.context} />
                                    </PopUp>

                                    <ThemeProvider theme={theme}>
                                        <TableRepeatingSection
                                            schema={employeeSchema}
                                            label={t.Attendees}
                                            name={'attendees'}
                                            data={attendees}
                                            addRow={() => addRow('attendees', 'Employee', this.setState.bind(this))}
                                            deleteRow={(rowIndex) => deleteRow('attendees', rowIndex, this.setState.bind(this))}
                                            onChangeGeneric={this.onChangeGeneric}
                                            context={this.props.context}
                                            companies={this.state.companies}
                                            users={users}
                                            currDir={currDir}
                                        />
                                        <TableRepeatingSection
                                            schema={employeeSchema}
                                            label={t.Absents}
                                            name={'absents'}
                                            data={absents}
                                            addRow={() => addRow('absents', 'Employee', this.setState.bind(this))}
                                            deleteRow={(rowIndex) => deleteRow('absents', rowIndex, this.setState.bind(this))}
                                            onChangeGeneric={this.onChangeGeneric}
                                            context={this.props.context}
                                            companies={this.state.companies}
                                            users={users}
                                            currDir={currDir}
                                        />

                                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'left', gap: '1em' }}>
                                            <Typography sx={{ paddingBottom: 0, display: 'flex' }} variant="h6" color="initial">{t.Copy}:</Typography>
                                            <div className={styles.copyContainer}>
                                                {[...this.state.selectedUsers, ...this.state.selectedUsersFreeSolo].flat().map((item: string, index: number) => (
                                                    <div key={uuidv4()} className={styles.copyItem}>
                                                        {item}
                                                        {this.state.selectedUsersFreeSolo.includes(item) && (
                                                            <IconButton
                                                                size='small' sx={{ display: 'flex', justifyContent: 'center' }} onClick={(e: any) => { this.deleteFreeSoloUser(e, item) }}>
                                                                <CloseIcon fontSize='small' />
                                                            </IconButton>
                                                        )}
                                                    </div>
                                                ))}
                                            </div>
                                        </div>

                                        <div style={{ display: 'flex', flexDirection: 'row', marginTop: '1em', marginBottom: '1em' }}>
                                            <TextField variant='standard' name='freeSoloUser' onBlur={this.onChange} type='text' />
                                            <Button onClick={this.onClickFreeSolo}><span style={{ fontSize: '20px' }}>+</span></Button>
                                        </div>
                                        <TableRepeatingSection
                                            schema={meetingContentSchema}
                                            label={t.meetingContent}
                                            name={'meetingContent'}
                                            data={meetingContent}
                                            addRow={() => addRow('meetingContent', 'MeetingContent', this.setState.bind(this))}
                                            deleteRow={(rowIndex) => deleteRow('meetingContent', rowIndex, this.setState.bind(this))}
                                            attachRow={(rowIndex) => this.attachRow('meetingContent', rowIndex)}
                                            onChangeGeneric={this.onChangeGeneric}
                                            context={this.props.context}
                                            companies={this.state.companies}
                                            users={users}
                                            currDir={currDir}
                                        />
                                        <TableRepeatingSection
                                            schema={taskSchema}
                                            label={t.Tasks}
                                            name={'tasks'}
                                            data={tasks}
                                            addRow={() => addRow('tasks', 'Task', this.setState.bind(this))}
                                            deleteRow={(rowIndex) => deleteRow('tasks', rowIndex, this.setState.bind(this))}
                                            onChangeGeneric={this.onChangeGeneric}
                                            context={this.props.context}
                                            companies={this.state.companies}
                                            users={users}
                                            currDir={currDir}
                                        />
                                    </ThemeProvider>

                                    <Divider style={{ paddingTop: '1em' }} />

                                    <div style={{ width: '100%', paddingTop: '2em' }}>

                                        <div style={{ display: 'flex', gap: '1em', alignItems: 'center' }}>
                                            <span>{t["Choose where to file the meeting summary"]}</span>
                                            <IconButton
                                                size='small'
                                                sx={{ display: 'flex', justifyContent: 'center' }} onClick={(e: any) => { this.setState({ folderPopUp: true }) }}>
                                                <SourceIcon fontSize='small' />
                                            </IconButton>
                                        </div>

                                        <PopUp open={this.state.folderPopUp} onClose={() => { this.closeFolderPopUp() }} title={t["Choose where to file the meeting summary"]} actions={null} dir={currDir ? 'rtl' : 'ltr'}>
                                            <Labeling selectedLabeling={this.state.selectedLabeling} sp={this.props.sp} context={this.props.context} dir={currDir} users={users} onSave={this.libraryPathHandle} onClose={this.closeFolderPopUp}></Labeling>
                                        </PopUp>

                                        <div style={{ display: 'flex', gap: '1em', alignItems: 'center' }}>
                                            <span>{t["File location"]}: </span>
                                            {libraryName !== '' ?
                                                <div style={{ display: 'flex', flexDirection: 'row', gap: '1em', alignItems: 'center' }}>
                                                    <span style={{ color: '#5989f6' }}>{libraryName}</span> <ClearIcon fontSize='small' className={styles.ClearIcon} onClick={() => { this.setState({ libraryName: '', libraryPath: '' }) }} />
                                                </div> : null}
                                        </div>
                                        {this.requireErrorJSX('libraryPath')}
                                    </div>

                                    {LoadingForm === 'Saving' ? <LinearProgress /> : null}

                                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', paddingTop: '2em', gap: '20px' }}>
                                        <Button disabled={this.state.submit === 'send'} variant="contained" color='success' sx={{ backgroundColor: '#8AC693', minWidth: '10em', textTransform: 'capitalize' }} onClick={() => this.submitForm('send')}>{t.SaveAndSend}</Button>
                                        <Button disabled={this.state.submit === 'send'} variant="contained" color='primary' sx={{ minWidth: '10em', textTransform: 'capitalize' }} onClick={() => this.submitForm('save')}>{t.Save}</Button>
                                        <Button variant="contained" color='error' sx={{ backgroundColor: '#CA3935', minWidth: '10em', textTransform: 'capitalize' }} onClick={() => sweetAlertMsgHandler('Cancel', currDir)}>{t.Cancel}</Button>
                                    </div>
                                </section>}
                            </div>
                        </Paper>
                    </form>
                </CacheProviderWrapper>

            </LocalizationProvider>
        );
    }
}