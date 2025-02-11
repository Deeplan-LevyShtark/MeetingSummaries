import * as React from "react";
import styles from "./NewContact.module.scss";
import { Autocomplete, Button, TextField } from "@mui/material";
import { SPFI } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INewContactProps {
    onClose: () => void;
    dir: boolean;
    sp: SPFI;
    context: WebPartContext;
}

export function NewContact(props: INewContactProps) {
    const [fullName, setFullName] = React.useState<string>("");
    const [email, setEmail] = React.useState<string>("");
    const [Options, setOptions] = React.useState<Array<string>>([]);
    const [users, setUsers] = React.useState<any>([]);
    const [company, setCompany] = React.useState<string>("");

    const [saving, setSaving] = React.useState<boolean>(false); // Track loading state

    React.useEffect(() => {
        async function fetchCompanies() {
            try {
                const users = await props.sp.web.lists.getByTitle("External Users Options").select('Title, Email').items();
                const companies = await props.sp.web.lists.getByTitle("Companies").items();
                const titles = companies.map((item: any) => item.Title);
                setOptions(titles);
                setUsers(users)
            } catch (error) {
                console.error("Error fetching companies:", error);
            }
        }

        fetchCompanies();
    }, [props.sp]);
    // Email validation regex
    function isValidEmail(email: string): boolean {
        const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
        return emailRegex.test(email);
    }

    // Name validation (at least 2 characters)
    function isValidFullName(name: string): boolean {
        return name?.trim().length >= 2;
    }

    // Save contact to SharePoint list
    async function saveToSP(): Promise<void> {
        if (!isValidEmail(email) || !isValidFullName(fullName)) return;

        setSaving(true); // Start loading state

        try {
            // Check if the user already exists
            const existingUser = users.find((user: any) => user.Email === email);
            // Check if the company already exists
            const existingCompany = Options.find((c: any) => c === company);

            let message = "";

            // Add user only if it does not exist
            if (existingUser) {
                message += props.dir
                    ? "האימייל כבר קיים במערכת. "
                    : "The email already exists in the system. ";
            }

            // Add company only if it does not exist
            if (existingCompany) {
                message += props.dir
                    ? "החברה כבר קיימת במערכת. "
                    : "The company already exists in the system. ";
            }

            // Execute all required actions in parallel
            if (existingCompany || existingUser) {
                alert(message);
            }

            if (!existingCompany && !existingUser) {
                await Promise.all([
                    props.sp.web.lists
                        .getByTitle("External Users Options")
                        .items.add({
                            Title: fullName, // Save Name
                            Email: email, // Save Email
                            Company: company
                        }),
                    props.sp.web.lists
                        .getByTitle("Companies")
                        .items.add({
                            Title: company
                        })
                ]);
                alert(props.dir ? "איש קשר נשמר בהצלחה!" : "Contact saved successfully!");
                props.onClose(); // Close after saving
            }

        } catch (error) {
            console.error("Error saving contact:", error);
            alert(props.dir ? "שמירת איש הקשר נכשלה" : "Failed to save contact.");
        } finally {
            setSaving(false); // End loading state
        }
    }

    return (
        <>
            <div className={styles.newContactContainer}>
                <TextField
                    size="small"
                    fullWidth
                    label={props.dir ? "שם מלא" : "Full Name"}
                    value={fullName}
                    onChange={(e) => setFullName(e.target.value)}
                    error={!isValidFullName(fullName) && fullName.length > 0}
                    helperText={!isValidFullName(fullName) && fullName.length > 0 ? props.dir ? "שם מלא חייב להכיל לפחות 2 אותיות" : "Full Name must be at least 2 characters" : ""}
                    required
                />
                <Autocomplete
                    freeSolo={true}
                    fullWidth
                    size="small"
                    options={Options}
                    //getOptionLabel={(option) => || ""}
                    value={company || null} // Controlled value
                    onChange={(event, newValue: any) =>
                        setCompany(newValue)
                    }
                    onInputChange={(event, newValue: any) => {
                        setCompany(newValue)
                    }}
                    renderInput={(params) => (
                        <TextField {...params} label={props.dir ? "חברה" : "Company"} variant="outlined" />
                    )}
                />
                <TextField
                    size="small"
                    fullWidth
                    label={props.dir ? "אימייל" : "Email"}
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    error={!isValidEmail(email) && email.length > 0}
                    helperText={!isValidEmail(email) && email.length > 0 ? props.dir ? "אימייל לא תקין" : "Invalid email address" : ""}
                    required
                />
            </div>

            <div style={{ display: "flex", justifyContent: "center", marginTop: "1rem", gap: "1rem" }}>
                <Button
                    style={{ textTransform: "capitalize" }}
                    size="small"
                    variant="contained"
                    onClick={saveToSP}
                    disabled={!isValidFullName(fullName) || !isValidEmail(email) || saving} // Disable if invalid or saving
                >
                    {saving ? (props.dir ? "שומר..." : "Saving...") : props.dir ? "שמור" : "Save"}
                </Button>

                <Button
                    style={{ textTransform: "capitalize" }}
                    size="small"
                    variant="contained"
                    color="error"
                    onClick={props.onClose}
                    disabled={saving} // Prevent closing while saving
                >
                    {props.dir ? "בטל" : "Cancel"}
                </Button>
            </div>
        </>
    );
}
