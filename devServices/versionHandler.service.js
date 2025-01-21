const fs = require('fs');

const filePath = './config/package-solution.json'; // Update with actual JSON file path

function incrementVersion(version) {
    const parts = version.split('.').map(part => parseInt(part, 10));

    for (let i = parts.length - 1; i >= 0; i--) {
        if (parts[i] < 9) {
            parts[i]++;
            break;
        } else {
            if (i === 0) {
                // If the first part is also 9, reset and add a new part at the beginning
                parts[i] = 0;
                parts.unshift(1);
                break;
            } else {
                parts[i] = 0;
            }
        }
    }

    return parts.join('.');
}

function updateVersion() {
    try {
        const json = JSON.parse(fs.readFileSync(filePath, 'utf8'));
        console.log("Current version:", json.solution.version);

        json.solution.version = incrementVersion(json.solution.version);
        fs.writeFileSync(filePath, JSON.stringify(json, json.solution.version, 2), 'utf8');

        console.log(`Updated version: ${json.solution.version}`);
    } catch (error) {
        console.error("Error updating version:", error);
    }
}

updateVersion();