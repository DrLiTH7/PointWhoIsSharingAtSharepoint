const msal = require('@azure/msal-node');
const axios = require('axios');
const {CLIENT_SECRET, SHAREPOINT_SITE, SHAREPOINT_DRIVE, GRAPH_SCOPE} = require('../config/config')
const msalConfig = require('../config/msalConfig')

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getAccessToken() {
    const authResult = await cca.acquireTokenByClientCredential({ scopes: GRAPH_SCOPE });
    return authResult.accessToken;
}

async function getFilesAndFolders(siteId, driveId, accessToken) {
    const headers = { Authorization: `Bearer ${accessToken}` };
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children`;

    try {
        const response = await axios.get(url, { headers });
        return response.data;
    } catch (error) {
        throw new Error(`Error fetching files and folders: ${error.message}`);
    }
}

async function getPermissionsForItem(siteId, driveId, itemId, accessToken) {
    const headers = { Authorization: `Bearer ${accessToken}` };
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/permissions`;

    try {
        const response = await axios.get(url, { headers });
        return response.data;
    } catch (error) {
        throw new Error(`Error fetching permissions: ${error.message}`);
    }
}

async function main() {
    try {
        // Step 1: Get access token
        const accessToken = await getAccessToken();

        // Step 2: Get files and folders from SharePoint
        const filesAndFolders = await getFilesAndFolders(SHAREPOINT_SITE, SHAREPOINT_DRIVE, accessToken);
        console.log("Files and Folders:");
        for (const item of filesAndFolders.value) {
            console.log(`Name: ${item.name}, ID: ${item.id}`);

            // Step 3: Get permissions for each file/folder
            const permissions = await getPermissionsForItem(SHAREPOINT_SITE, SHAREPOINT_DRIVE, item.id, accessToken);
            console.log("Permissions:");
            for (const perm of permissions.value) {
                const grantedTo = perm.grantedTo?.user?.displayName || "N/A";
                const roles = perm.roles.join(', ');
                console.log(`Granted To: ${grantedTo}, Roles: ${roles}`);
            }
            console.log("\n");
        }

    } catch (error) {
        console.error(`Error: ${error.message}`);
    }
}

main();
