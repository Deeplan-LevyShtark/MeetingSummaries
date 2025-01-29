import * as React from "react";
import styles from "./NewContact.module.scss";
import { Button, TextField } from "@mui/material";
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
    const [saving, setSaving] = React.useState<boolean>(false); // Track loading state

    // Email validation regex
    function isValidEmail(email: string): boolean {
        const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
        return emailRegex.test(email);
    }

    // Name validation (at least 2 characters)
    function isValidFullName(name: string): boolean {
        return name.trim().length >= 2;
    }

    // Save contact to SharePoint list
    async function saveToSP(): Promise<void> {
        if (!isValidEmail(email) || !isValidFullName(fullName)) return;

        setSaving(true); // Start loading state

        try {
            await props.sp.web.lists
                .getByTitle("External Users Options") // Ensure list exists
                .items.add({
                    Title: fullName, // Save Name
                    Email: email, // Save Email
                });

            alert(props.dir ? "איש קשר נשמר בהצלחה" : "Contact saved successfully!");
            props.onClose(); // Close after saving
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
