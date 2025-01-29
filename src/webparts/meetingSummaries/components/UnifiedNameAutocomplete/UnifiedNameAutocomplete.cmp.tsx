import * as React from "react";
import Autocomplete from "@mui/material/Autocomplete";
import TextField from "@mui/material/TextField";
import { Avatar, Tooltip, Chip } from "@mui/material";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { v4 as uuidv4 } from "uuid";
import CancelIcon from '@mui/icons-material/Cancel';
import { update } from "@microsoft/sp-lodash-subset";

interface User {
    Id: number;
    Title: string;
    Email: string;
}

interface UnifiedNameAutocompleteProps {
    users: User[];
    context: WebPartContext;
    value?: string | string[];
    label?: string;
    sx?: {};
    params?: { id: number };
    multiple?: boolean;
    freeSolo?: boolean;
    size?: "small" | "medium";
    onChange: (
        idOrValue: number | string | (number | string)[],
        newValue: string | string[],
        email?: string | string[],
        isFreeSolo?: boolean,
        ids?: string[],
        onBlur?: string
    ) => void;
}

export const UnifiedNameAutocomplete: React.FC<UnifiedNameAutocompleteProps> = ({
    users,
    context,
    value = "",
    label,
    sx,
    params,
    multiple = false,
    freeSolo = false,
    size = "medium",
    onChange,
}) => {
    const [localInputValue, setLocalInputValue] = React.useState<string | string[]>(value || "");
    const [internalValue, setInternalValue] = React.useState<string | string[]>(value || "");
    const textFieldRef = React.useRef<HTMLInputElement>(null);

    React.useEffect(() => {
        if (Array.isArray(value)) {
            setLocalInputValue(value.map((v: User | string) => (typeof v === "string" ? v : v.Title)));
            setInternalValue(value.map((v) => (typeof v === "string" ? v : (v as User).Title)));
        } else {
            setLocalInputValue(value || "");
            setInternalValue(value || "");
        }
    }, [value]);


    const handleBlur = () => {

        if (Array.isArray(localInputValue)) {
            const ids = localInputValue.map((title) => {
                const user = users.find((u) => u.Title === title);
                return user ? user.Id.toString() : "";
            }).filter(id => id !== ""); // Filter out empty strings
            // i want to filter the emails from the user by title
            const filterdEmails = localInputValue.map((title) => {
                const user = users.find((u) => u.Title === title);
                return user ? user.Email : "";
            }).filter(email => email !== ""); // Filter out empty strings      
            onChange(params?.id ?? localInputValue, localInputValue, filterdEmails, true, ids, 'onBlur');
        }
        else if (localInputValue.trim() !== "" && (label === 'Name' || label !== 'ForInfo')) {
            const updatedValue = [...(internalValue as string[]), localInputValue];
            setInternalValue(updatedValue);
            const filterdEmails = users.filter((u) => (updatedValue as string[]).includes(u.Title)).map((u) => u.Email);
            onChange(params?.id ?? updatedValue, updatedValue, filterdEmails, true);
        }
    };

    const handleInputChange = (event: any, newInputValue: string) => {
        setLocalInputValue(newInputValue);
    };

    const handleOptionChange = (
        event: React.ChangeEvent<{}>,
        newValue: User | string | (User | string)[] | null
    ) => {
        if (multiple) {
            if (Array.isArray(newValue)) {

                const updatedValue = Array.from(new Set([...(internalValue as string[]), ...newValue.map((item) => (typeof item === "string" ? item : item.Title))]));

                const selectedIds = newValue.map((item) => (typeof item === "string" ? "" : item.Id.toString()));
                setInternalValue(updatedValue); // Update state with selected Titles

                const selectedEmails = newValue.map((item) => (typeof item === "string" ? "" : item.Email));

                onChange(params?.id ?? updatedValue, updatedValue, selectedEmails, true, selectedIds); // Pass to parent component                

            } else if (newValue && typeof newValue === "string") {
                const updatedValue = [...(internalValue as string[]), newValue];
                const filterdEmails = users.filter((u) => u.Title === newValue).map((u) => u.Email);
                setInternalValue(updatedValue);
                onChange(params?.id ?? updatedValue, updatedValue, filterdEmails, true);
            }
        } else {
            if (typeof newValue === "string") {
                setInternalValue(newValue);
                onChange(params?.id ?? newValue, newValue, "", true);
            } else if (newValue && "Title" in newValue) {
                setInternalValue(newValue.Title);
                onChange(params?.id ?? newValue.Id, newValue.Title, newValue.Email, false, [newValue.Id.toString()]);
            } else {
                setInternalValue("");
                onChange(params?.id ?? "", "", "", false);
            }
        }
    };

    const handleDelete = (titleToDelete: string) => {
        if (textFieldRef.current) {
            textFieldRef.current.focus();
        }

        const updatedLocalInputValue = (localInputValue as string[]).filter((title) => title !== titleToDelete);
        const updatedInternalValue = (internalValue as string[]).filter((title) => title !== titleToDelete);

        setLocalInputValue(updatedLocalInputValue); // Update local state
        setInternalValue(updatedInternalValue); // Update internal state

    };

    return (
        <Autocomplete
            size={size}
            sx={sx}
            options={users}
            disableClearable
            getOptionLabel={(option) => {
                if (typeof option === "string") return option;
                if (option?.Title) return option.Title;
                return "";
            }}
            value={
                multiple
                    ? users.filter((u) => (Array.isArray(internalValue) ? internalValue : []).includes(u.Title))
                    : users.find((u) => u.Title === internalValue) || null
            }
            freeSolo={freeSolo}
            multiple={multiple}
            inputValue={typeof localInputValue === "string" ? localInputValue : ""}
            onInputChange={handleInputChange}
            onChange={handleOptionChange}
            onBlur={handleBlur}
            isOptionEqualToValue={(option, value) => {
                if (typeof value === "string") {
                    return option.Title === value;
                }
                return option.Title === value.Title;
            }}
            renderTags={(value: string[], getTagProps) => {

                return value.map((option: any, index) => (
                    <Chip
                        label={option.Title}
                        {...getTagProps({ index })}
                        // onDelete={() => handleDelete(option.Title)}
                        deleteIcon={<CancelIcon />}
                        style={{ marginRight: 8 }}
                    />
                ));
            }}
            renderInput={(inputParams) => (
                <Tooltip
                    title={
                        Array.isArray(localInputValue)
                            ? localInputValue.join(", ") // Show all selected Titles as a tooltip
                            : localInputValue
                    }
                >
                    <TextField
                        {...inputParams}
                        // label={label}
                        onBlur={handleBlur}
                        inputRef={textFieldRef}
                        InputProps={{
                            ...inputParams.InputProps,
                            startAdornment:
                                multiple && Array.isArray(localInputValue) && (localInputValue as string[]).map((title) => (
                                    <Chip
                                        key={uuidv4()}
                                        avatar={
                                            <Avatar
                                                src={`${context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?UserName=${users.find(
                                                    (u) => u.Title === title
                                                )?.Email}&size=L`}
                                                alt={title}
                                            />
                                        }
                                        label={title}
                                        onDelete={(e) => {
                                            e.preventDefault();
                                            handleDelete(title);
                                        }}
                                        deleteIcon={<CancelIcon />}
                                        style={{ marginRight: 8 }}
                                    />
                                ))

                            , style: { display: "flex" }
                        }}
                    />
                </Tooltip>
            )}

            renderOption={(props: any, option: any) => (
                <li
                    {...props}
                    key={uuidv4()}
                    style={{ display: "flex", alignItems: "center", width: "100%" }}
                >
                    <Avatar
                        src={`${context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?UserName=${option.Email}&size=L`}
                        alt={option.Title}
                        style={{ marginRight: 8, width: 30, height: 30 }}
                    />
                    <div style={{ display: "flex", flexDirection: "column" }}>
                        {option.Title}
                        <span style={{ fontSize: 12, color: "gray" }}>{option.Email}</span>
                    </div>
                </li>
            )}
            fullWidth
        />
    );
};