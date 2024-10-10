const {TENANT_ID, CLIENT_ID} = require('config')

module.exports = msalConfig = {
    auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        clientSecret: CLIENT_SECRET,
    }
};