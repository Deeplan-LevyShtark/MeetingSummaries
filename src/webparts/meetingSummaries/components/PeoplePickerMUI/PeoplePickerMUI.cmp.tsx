import * as React from "react";
import Autocomplete from "@mui/material/Autocomplete";
import { TextField, Chip, InputAdornment } from "@mui/material";
import Avatar from "@mui/material/Avatar";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { v4 as uuidv4 } from "uuid";

interface PeoplePickerMUIProps {
    context: WebPartContext;
    users: any[];
    multiple: boolean;
    selectedUser: any;
    onChange: (value: any) => void;
    name?: string;
    label?: string;
    error?: boolean;
}

export default function PeoplePickerMUI({ context, users, multiple, selectedUser, name, label, error, onChange, }: PeoplePickerMUIProps) {

    // Handle onDelete to remove a user
    const handleDelete = () => {
        onChange(null);
    };

    return (
        <Autocomplete
            onChange={(event, newValue) => {
                onChange(newValue || null);
            }}
            fullWidth
            value={selectedUser}
            multiple={false}
            size="medium"
            options={users}

            getOptionLabel={(option) => {
                if (typeof option === "string") return option;
                if (option?.Title) return option.Title;
                return "";
            }}
            isOptionEqualToValue={(option, value) => option.Email === value.Email}
            renderOption={(props, option) => (
                <li
                    {...props}
                    style={{ display: "flex", alignItems: "center", width: "100%" }}
                    key={uuidv4()}
                >
                    <Avatar

                        src={`${context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?UserName=${option.Email}&size=L`}
                        alt={option.Title}
                        style={{ marginRight: 8, width: 30, height: 30 }}
                    />
                    <div
                        style={{ display: "flex", flexDirection: "column", width: "100%" }}
                    >
                        {option.Title}
                        <span style={{ fontSize: 12, color: "gray" }}>{option.Email}</span>
                    </div>
                </li>
            )}
            renderInput={(params) => (
                <TextField

                    {...params}
                    error={error}
                    label={label}
                    InputProps={{
                        ...params.InputProps,

                    }}
                > </TextField>
            )}
        />
    );
}